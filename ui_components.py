import base64
import html
from pathlib import Path
import streamlit as st

from file_utils import load_prefs, save_prefs

# Uniform Lucide 24x24 SVG icon set (stroke-width="2", round caps & joins)
_ICON_LAYERS = (
    '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" '
    'stroke-linecap="round" stroke-linejoin="round">'
    '<path d="m12.83 2.18a2 2 0 0 0-1.66 0L2.6 6.08a1 1 0 0 0 0 1.83l8.58 3.91a2 2 0 0 0 1.66 0l8.58-3.9a1 1 0 0 0 0-1.83Z"/>'
    '<path d="m22 17.65-9.17 4.16a2 2 0 0 1-1.66 0L2 17.65"/>'
    '<path d="m22 12.65-9.17 4.16a2 2 0 0 1-1.66 0L2 12.65"/></svg>'
)
_ICON_SPARK = (
    '<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" '
    'stroke-linecap="round" stroke-linejoin="round">'
    '<path d="M9.937 15.5A2 2 0 0 0 8.5 14.063l-6.135-1.582a.5.5 0 0 1 0-.962L8.5 9.936A2 2 0 0 0 9.937 8.5l1.582-6.135a.5.5 0 0 1 .963 0L14.063 8.5A2 2 0 0 0 15.5 9.937l6.135 1.581a.5.5 0 0 1 0 .964L15.5 14.063a2 2 0 0 0-1.437 1.437l-1.582 6.135a.5.5 0 0 1-.963 0z"/>'
    '<path d="M20 3v4"/><path d="M22 5h-4"/><path d="M4 17v2"/><path d="M5 18H3"/></svg>'
)

# File icon on the right of the "Save settings" card heading (Lucide file-text)
_ICON_FILE = (
    '<svg aria-hidden="true" width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="#f97316" '
    'stroke-width="2" stroke-linecap="round" stroke-linejoin="round">'
    '<path d="M15 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V7Z"/>'
    '<path d="M14 2v4a2 2 0 0 0 2 2h4"/><path d="M10 9H8"/><path d="M16 13H8"/><path d="M16 17H8"/></svg>'
)

# The "?" icon of the Help link in the header (Lucide circle-help)
_ICON_HELP = (
    '<svg aria-hidden="true" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" '
    'stroke-width="2" stroke-linecap="round" stroke-linejoin="round">'
    '<circle cx="12" cy="12" r="10"/>'
    '<path d="M9.09 9a3 3 0 0 1 5.83 1c0 2-3 3-3 3"/><path d="M12 17h.01"/></svg>'
)

# Sun & Moon Lucide 24x24 SVG icons for Dark Mode reference
_ICON_SUN = (
    '<svg aria-hidden="true" width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" '
    'stroke-width="2" stroke-linecap="round" stroke-linejoin="round">'
    '<circle cx="12" cy="12" r="4"/><path d="M12 2v2"/><path d="M12 20v2"/>'
    '<path d="m4.93 4.93 1.41 1.41"/><path d="m17.66 17.66 1.41 1.41"/>'
    '<path d="M2 12h2"/><path d="M20 12h2"/>'
    '<path d="m6.34 17.66-1.41 1.41"/><path d="m19.07 4.93-1.41 1.41"/></svg>'
)
_ICON_MOON = (
    '<svg aria-hidden="true" width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" '
    'stroke-width="2" stroke-linecap="round" stroke-linejoin="round">'
    '<path d="M12 3a6 6 0 0 0 9 9 9 9 0 1 1-9-9Z"/></svg>'
)

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

_DARK_THEME_OVERRIDE_CSS = """
:root {
  --brand-navy: #2547a8;
  --brand-navy-mid: #1e3a8a;
  --brand-navy-dark: #0b1536;
  --brand-navy-deep: #060d21;
  --brand-navy-soft: #1e293b;
  --brand-navy-tint: #172554;
  --brand-orange: #f97316;
  --brand-orange-bright: #fb923c;
  --brand-orange-hover: #ea580c;
  --brand-orange-soft: rgba(249, 115, 22, 0.18);
  --brand-orange-tint: rgba(234, 88, 12, 0.12);
  --status-success: #22c55e;
  --status-success-bg: rgba(22, 163, 74, 0.18);
  --status-success-border: rgba(34, 197, 94, 0.4);
  --status-warn: #fb923c;
  --status-warn-bg: rgba(234, 88, 12, 0.18);
  --status-warn-border: rgba(251, 146, 60, 0.45);
  --status-danger: #f87171;
  --status-danger-bg: rgba(220, 38, 38, 0.18);
  --status-danger-border: rgba(248, 113, 113, 0.45);
  --bg-canvas: linear-gradient(135deg, #060d21 0%, #0f172a 50%, #431407 88%, #7c2d12 100%);
  --surface-header: linear-gradient(90deg, #060d21 0%, #0b1536 55%, #112252 100%);
  --surface-shell: #0f172a;
  --surface-panel-primary: #0f172a;
  --surface-panel-secondary: #111c35;
  --surface-sidebar: linear-gradient(180deg, #080f26 0%, #0d183a 100%);
  --surface-card: #1e293b;
  --surface-elevated: #162033;
  --surface-input: #0f172a;
  --surface-input-hover: #162238;
  --surface-table-row-alt: #162033;
  --surface-table-hover: #1e2d4a;
  --surface-group-row: #172554;
  --text-primary: #f8fafc;
  --text-secondary: #e2e8f0;
  --text-body: #cbd5e1;
  --text-muted: #94a3b8;
  --text-on-navy: #f8fafc;
  --text-on-navy-muted: #cbd5e1;
  --border-subtle: #334155;
  --border-strong: #475569;
  --border-navy: #1e3a8a;
  --border-shell: rgba(148, 163, 184, 0.18);
  --shadow-sm: 0 2px 4px rgba(0, 0, 0, 0.35), 0 4px 12px rgba(0, 0, 0, 0.25);
  --shadow-md: 0 6px 14px -2px rgba(0, 0, 0, 0.45), 0 12px 28px -4px rgba(0, 0, 0, 0.35);
  --shadow-lg: 0 16px 36px -4px rgba(0, 0, 0, 0.55), 0 6px 16px -2px rgba(0, 0, 0, 0.40);
}
"""


def _ensure_dark_mode_state():
    """Initializes st.session_state['dark_mode'] from saved user preferences if not yet loaded."""
    if "dark_mode" not in st.session_state:
        st.session_state["dark_mode"] = bool(load_prefs().get("dark_mode", False))
    return bool(st.session_state.get("dark_mode", False))


def _toggle_dark_mode():
    """Flips dark_mode in session_state and persists it to user_prefs.json."""
    new_val = not bool(st.session_state.get("dark_mode", False))
    st.session_state["dark_mode"] = new_val
    save_prefs({"dark_mode": new_val})


def _compact_css(raw_css):
    """Strips blank lines and leading/trailing spaces so CommonMark never parses indented
    CSS lines as a <pre><code> block or terminates HTML blocks early."""
    return " ".join(line.strip() for line in raw_css.splitlines() if line.strip())


def load_local_css(file_name="style.css"):
    """Injects the external CSS file into the Streamlit app layout and applies active theme tokens
    in a single compact <style> block."""
    dark_mode = _ensure_dark_mode_state()
    css_path = Path(__file__).with_name(file_name)
    if not css_path.exists():
        css_path = Path(file_name)
    css_parts = []
    try:
        css = css_path.read_text(encoding="utf-8").replace("\u00a0", " ")
        css_parts.append(_compact_css(css))
    except OSError:
        pass
    if dark_mode:
        css_parts.append(_compact_css(_DARK_THEME_OVERRIDE_CSS))
    if css_parts:
        st.markdown(f"<style>{' '.join(css_parts)}</style>", unsafe_allow_html=True)


def _brand_html(dark_mode=False):
    """The logo, a thin orange divider and the product name (left of the header bar)."""
    logo_path = Path(__file__).with_name("vonotec.png")
    if not logo_path.exists():
        logo_path = Path("vonotec.png")
    try:
        logo_base64 = base64.b64encode(logo_path.read_bytes()).decode()
        logo = f'<img src="data:image/png;base64,{logo_base64}" width="132" alt="VONOTEC">'
    except OSError:
        logo = '<span class="brand-text">VONOTEC</span>'
    theme_marker = '<span data-theme="dark" style="display:none;"></span>' if dark_mode else ""
    return f'<div class="brand">{theme_marker}{logo}<span class="product-name">Automated Defect Logger</span></div>'


def render_header(views):
    """Deep Vonotec Navy header card with three columns: the brand on the left,
    the view switch (a radio, drawn as two pills) in the middle, and the
    tagline + Help popover + Dark Mode toggle on the right."""
    dark_mode = _ensure_dark_mode_state()

    with st.container(key="top_app_header"):
        try:
            col_left, col_center, col_right, col_theme = st.columns([1.0, 0.76, 1.14, 0.10], vertical_alignment="center")
        except TypeError:
            col_left, col_center, col_right, col_theme = st.columns([1.0, 0.76, 1.14, 0.10])

        with col_left:
            st.markdown(_brand_html(dark_mode=dark_mode), unsafe_allow_html=True)
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
                '<span class="header-tagline">Defect Extraction &amp; Repair Analytics</span>'
                f'<details class="help-pop"><summary>{_ICON_HELP}Help</summary>'
                f'<div class="help-panel">{_HELP_HTML}</div></details>'
                '</div>',
                unsafe_allow_html=True,
            )
        with col_theme:
            toggle_key = "theme_toggle_light" if dark_mode else "theme_toggle_dark"
            with st.container(key=toggle_key):
                st.button(
                    " ",
                    key="theme_toggle_btn",
                    on_click=_toggle_dark_mode,
                    help="Switch to Light mode" if dark_mode else "Switch to Dark mode",
                )
    return view


WORKFLOW_STEPS = ("Select photos", "Save destination", "Extract data")


def render_stepper(done, current):
    """A compact, transparent "1. Select photos > 2. Save destination > 3. Extract data" line."""
    parts = []
    for number, label in enumerate(WORKFLOW_STEPS, start=1):
        state = "done" if number in done else ("current" if number == current else "todo")
        mark = "&#10003;" if state == "done" else str(number)
        parts.append(f'<span class="wf-step wf-{state}"><span class="wf-num">{mark}</span>{html.escape(label)}</span>')
    st.markdown(
        '<div class="workflow-stepper">' + '<span class="wf-arrow">&rarr;</span>'.join(parts) + '</div>',
        unsafe_allow_html=True,
    )


def section_heading(eyebrow, title, badge=None, step=None, description=None):
    """A small orange label over a large title. On the right there is either a
    status badge, badge = (text, kind) with kind "ready", "processing" or
    "complete", or a plain step label."""
    right = ""
    if badge:
        text, kind = badge
        right = f'<span class="status-badge {html.escape(kind)}"><span></span>{html.escape(text)}</span>'
    elif step:
        right = f'<span class="step-label">{html.escape(step)}</span>'
    desc_html = f'<p class="section-desc">{html.escape(description)}</p>' if description else ""
    st.markdown(
        '<div class="section-heading"><div>'
        f'<p class="eyebrow">{html.escape(eyebrow)}</p>'
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


# Small icons for the "After extraction, you'll see" preview (Lucide circle-check, triangle-alert, file-spreadsheet)
_ICON_PREVIEW_STATUS = (
    '<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" '
    'stroke-linecap="round" stroke-linejoin="round">'
    '<circle cx="12" cy="12" r="10"/><path d="m9 12 2 2 4-4"/></svg>'
)
_ICON_PREVIEW_REVIEW = (
    '<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" '
    'stroke-linecap="round" stroke-linejoin="round">'
    '<path d="m21.73 18-8-14a2 2 0 0 0-3.48 0l-8 14A2 2 0 0 0 4 21h16a2 2 0 0 0 1.73-3"/>'
    '<path d="M12 9v4"/><path d="M12 17h.01"/></svg>'
)
_ICON_PREVIEW_FILE = (
    '<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" '
    'stroke-linecap="round" stroke-linejoin="round">'
    '<path d="M15 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V7Z"/>'
    '<path d="M14 2v4a2 2 0 0 0 2 2h4"/><path d="M8 13h2"/><path d="M14 13h2"/><path d="M8 17h2"/><path d="M14 17h2"/></svg>'
)


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
                Add whiteboard photos from your device or from a SharePoint folder. The AI will automatically extract the data
                and append it to the Master Excel file.
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
