# ui_components.py
import base64
import html
from pathlib import Path
import streamlit as st

_ICON_LAYERS = (
    '<svg width="23" height="23" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" '
    'stroke-linecap="round" stroke-linejoin="round"><path d="m12 3 9 5-9 5-9-5 9-5Zm-7.5 9 7.5 4 7.5-4M4.5 16l7.5 4 7.5-4"/></svg>'
)
_ICON_SPARK = (
    '<svg width="18" height="18" viewBox="0 0 24 24" fill="currentColor"><path d="m12 3 1.2 4.1L17 9l-3.8 1.9L12 15l-1.2-4.1L7 9l3.8-1.9L12 3ZM5 15l.7 2.3L8 18.5l-2.3 1.2L5 22l-.7-2.3L2 18.5l2.3-1.2L5 15Zm14-3 .7 2.3 2.3 1.2-2.3 1.2L19 19l-.7-2.3-2.3-1.2 2.3-1.2L19 12Z"/></svg>'
)
_ICON_FILE = (
    '<svg aria-hidden="true" width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="#9ba8b9" '
    'stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round">'
    '<path d="M7 3.5h6l4 4V20H7V3.5Zm6 0v4h4M9.5 12h5m-5 3h5"/></svg>'
)
_ICON_HELP = (
    '<svg aria-hidden="true" width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" '
    'stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="9"/>'
    '<path d="M9.5 9.2a2.6 2.6 0 0 1 5 .9c0 1.7-2.5 2.2-2.5 3.9M12 17h.01"/></svg>'
)

_HELP_HTML = (
    '<div class="help-title">How to use the whiteboard extractor</div>'
    '<p>Add whiteboard photos from your device or from a SharePoint folder. The AI extracts the data '
    'and appends it to the Master Excel file.</p>'
    '<ol>'
    '<li>Choose where to import from: <b>Device upload</b> or <b>Import from SharePoint</b>, then select your photos or folder.</li>'
    '<li>Once selected, <b>Save settings</b> unlocks in Step 2. Choose the SharePoint output folder and workbook name.</li>'
    '<li>Click the <b>Extract</b> button.</li>'
    '<li>Follow the progress, results and download on the right. The <b>Dashboard</b> picks up the new data automatically.</li>'
    '</ol>'
)

_ICON_PREVIEW_STATUS = (
    '<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" '
    'stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="9"/><path d="m8.5 12.3 2.4 2.4 4.6-5"/></svg>'
)
_ICON_PREVIEW_REVIEW = (
    '<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" '
    'stroke-linecap="round" stroke-linejoin="round"><path d="M12 4 2.8 19.5h18.4L12 4Z"/><path d="M12 10v4.2M12 17h.01"/></svg>'
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
    return f'<div class="brand">{logo}<span class="product-name">Whiteboard AI</span></div>'


def render_header(views):
    """One white header card with three equal columns."""
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
                '<span class="header-tagline">Data Extraction and Master Log Automator</span>'
                f'<details class="help-pop"><summary>{_ICON_HELP}Help</summary>'
                f'<div class="help-panel">{_HELP_HTML}</div></details>'
                '</div>',
                unsafe_allow_html=True,
            )
    return view


WORKFLOW_STEPS = ("Select photos", "Save destination", "Extract data")


def render_stepper(done, current=None):
    """A compact "1. Select photos > 2. Save destination > 3. Extract data" line."""
    if isinstance(done, int) and current is None:
        current = done
        done = set(range(1, done))
    else:
        done = set(done or ())
        current = current or 1

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
    """A small orange label over a large title, with an optional status badge on the right."""
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


def render_locked_save_settings():
    """Renders the compact locked 'Save settings' card when Step 1 is not yet complete."""
    st.markdown(
        '<div class="locked-save-card">'
        '  <div class="locked-save-left">'
        '    <div class="locked-save-icon">'
        '      <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="#64748b" stroke-width="1.9" stroke-linecap="round" stroke-linejoin="round">'
        '        <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>'
        '        <polyline points="14 2 14 8 20 8"/>'
        '        <line x1="16" y1="13" x2="8" y2="13"/>'
        '        <line x1="16" y1="17" x2="8" y2="17"/>'
        '      </svg>'
        '    </div>'
        '    <div class="locked-save-copy">'
        '      <div class="locked-save-title">Save settings</div>'
        '      <div class="locked-save-sub">Available after photos or a SharePoint folder are selected.</div>'
        '    </div>'
        '  </div>'
        '  <span class="locked-step-pill">Step 2</span>'
        '</div>',
        unsafe_allow_html=True,
    )


def render_empty_results():
    """Renders the consolidated 'No batch results yet' card and 3-outcome preview strip."""
    st.markdown(
        '<div class="empty-results-card">'
        '  <div class="empty-results-icon">'
        '    <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="#1e3a8a" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">'
        '      <polygon points="12 2 2 7 12 12 22 7 12 2"/>'
        '      <polyline points="2 12 12 17 22 12"/>'
        '      <polyline points="2 17 12 22 22 17"/>'
        '    </svg>'
        '  </div>'
        '  <div class="empty-results-body">'
        '    <div class="empty-results-title">No batch results yet</div>'
        '    <p class="empty-results-sub">Select whiteboard photos to see processing status, flagged rows, and the generated workbook.</p>'
        '  </div>'
        '</div>'
        '<div class="output-includes-box">'
        '  <div class="output-includes-label">BATCH OUTPUT INCLUDES</div>'
        '  <div class="output-includes-grid">'
        '    <div class="output-pill-item">'
        '      <span class="output-pill-icon blue">'
        '        <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="#1e3a8a" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round">'
        '          <circle cx="12" cy="12" r="10"/><polyline points="9 12 11 14 15 10"/>'
        '        </svg>'
        '      </span>'
        '      <span class="output-pill-text">Processing status</span>'
        '    </div>'
        '    <div class="output-pill-item">'
        '      <span class="output-pill-icon amber">'
        '        <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="#d97706" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round">'
        '          <path d="M12 3l1.9 5.8a2 2 0 0 0 1.3 1.3L21 12l-5.8 1.9a2 2 0 0 0-1.3 1.3L12 21l-1.9-5.8a2 2 0 0 0-1.3-1.3L3 12l5.8-1.9a2 2 0 0 0 1.3-1.3L12 3z"/>'
        '        </svg>'
        '      </span>'
        '      <span class="output-pill-text">Flagged records</span>'
        '    </div>'
        '    <div class="output-pill-item">'
        '      <span class="output-pill-icon blue">'
        '        <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="#1e3a8a" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round">'
        '          <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>'
        '          <polyline points="14 2 14 8 20 8"/>'
        '          <line x1="16" y1="13" x2="8" y2="13"/><line x1="16" y1="17" x2="8" y2="17"/>'
        '        </svg>'
        '      </span>'
        '      <span class="output-pill-text">Excel workbook</span>'
        '    </div>'
        '  </div>'
        '</div>',
        unsafe_allow_html=True,
    )


def render_tip(completed=False, review_count=0):
    """Contextual footer card in the results panel."""
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