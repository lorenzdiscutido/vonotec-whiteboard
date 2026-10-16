# ui_components.py
"""Shared visual building blocks: top header bar, 3-stage workflow stepper,
section/card headings, compact locked Save settings state, empty results preview,
and contextual extraction tips."""
import base64
import html
import streamlit as st


def load_local_css(file_name="style.css"):
    """Injects the external CSS file into the Streamlit app layout."""
    try:
        with open(file_name, encoding="utf-8") as f:
            st.markdown(f"<style>{f.read()}</style>", unsafe_allow_html=True)
    except FileNotFoundError:
        pass


_HELP_STEPS = (
    ("1. Select photos", "Choose <b>Device upload</b> to drop photos from your computer, or <b>Import from SharePoint</b> to pick a project folder."),
    ("2. Save destination", "Once photos are selected, Step 2 unlocks. Pick the SharePoint output folder and review the Excel workbook name."),
    ("3. Extract data", "Click <b>Extract photos</b>. Review any flagged items on the right and download the generated Master Log."),
)


def _help_popover_html():
    items = "".join(
        f'<div class="help-step"><div class="help-step-title">{title}</div>'
        f'<div class="help-step-body">{body}</div></div>'
        for title, body in _HELP_STEPS
    )
    return (
        '<details class="help-popover">'
        '<summary class="help-trigger">'
        '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" '
        'stroke-width="2" stroke-linecap="round" stroke-linejoin="round">'
        '<circle cx="12" cy="12" r="10"/>'
        '<path d="M9.09 9a3 3 0 0 1 5.83 1c0 2-3 3-3 3"/>'
        '<line x1="12" y1="17" x2="12.01" y2="17"/>'
        '</svg>'
        '<span>Help</span>'
        '</summary>'
        f'<div class="help-panel"><div class="help-panel-head">How Whiteboard AI works</div>{items}</div>'
        '</details>'
    )


def render_header(views):
    """Renders the 3-column enterprise top header card and returns the selected view."""
    logo_html = ""
    try:
        with open("vonotec.png", "rb") as f:
            logo_b64 = base64.b64encode(f.read()).decode()
        logo_html = f'<img src="data:image/png;base64,{logo_b64}" class="brand-logo" alt="VONOTEC">'
    except FileNotFoundError:
        logo_html = '<span class="brand-fallback">VONOTEC</span>'

    with st.container(key="top_app_header"):
        col_left, col_center, col_right = st.columns([1.15, 1.1, 1.15], vertical_alignment="center")
        with col_left:
            st.markdown(
                f'<div class="header-brand">'
                f'{logo_html}'
                f'<span class="brand-divider"></span>'
                f'<span class="brand-product">Whiteboard AI</span>'
                f'</div>',
                unsafe_allow_html=True,
            )
        with col_center:
            chosen = st.radio(
                "View",
                views,
                horizontal=True,
                label_visibility="collapsed",
                key="nav_view_switcher",
            )
        with col_right:
            st.markdown(
                f'<div class="header-right">'
                f'<span class="header-tagline">Data Extraction &amp; Master Log Automator</span>'
                f'{_help_popover_html()}'
                f'</div>',
                unsafe_allow_html=True,
            )
    return chosen


def render_stepper(done_steps, current_step):
    """Draws the 3-step horizontal progress stepper above the workspace."""
    steps = [
        (1, "Select photos"),
        (2, "Save destination"),
        (3, "Extract data"),
    ]
    parts = []
    for num, label in steps:
        if num in done_steps:
            state_cls = "done"
            badge_inner = "&#10003;"
        elif num == current_step:
            state_cls = "active"
            badge_inner = str(num)
        else:
            state_cls = "upcoming"
            badge_inner = str(num)

        parts.append(
            f'<div class="wf-step {state_cls}">'
            f'<span class="wf-num">{badge_inner}</span>'
            f'<span class="wf-label">{html.escape(label)}</span>'
            f'</div>'
        )
        if num < len(steps):
            line_cls = "done" if num in done_steps else ""
            parts.append(f'<div class="wf-line {line_cls}"></div>')

    st.markdown(f'<div class="wf-stepper">{"".join(parts)}</div>', unsafe_allow_html=True)


def section_heading(eyebrow, title, badge=None, description=""):
    """Renders an uppercase orange eyebrow, bold section title, optional status pill, and description."""
    badge_html = ""
    if badge:
        text, kind = badge
        badge_html = (
            f'<span class="sec-badge {html.escape(kind)}">'
            f'<span class="sec-badge-dot"></span>{html.escape(text)}</span>'
        )
    desc_html = f'<p class="sec-desc">{html.escape(description)}</p>' if description else ""
    st.markdown(
        f'<div class="section-heading">'
        f'<div class="sec-top-row">'
        f'<div><p class="eyebrow">{html.escape(eyebrow)}</p><h2>{html.escape(title)}</h2></div>'
        f'{badge_html}'
        f'</div>'
        f'{desc_html}'
        f'</div>',
        unsafe_allow_html=True,
    )


def subsection_heading(title, subtitle=""):
    sub_html = f'<p class="subsec-sub">{html.escape(subtitle)}</p>' if subtitle else ""
    st.markdown(
        f'<div class="subsection-heading">'
        f'<div class="subsec-title">{html.escape(title)}</div>'
        f'{sub_html}'
        f'</div>',
        unsafe_allow_html=True,
    )


def card_heading(eyebrow, title):
    st.markdown(
        f'<div class="card-heading">'
        f'<div><p class="eyebrow">{html.escape(eyebrow)}</p>'
        f'<div class="card-title">{html.escape(title)}</div></div>'
        f'</div>',
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
    """Contextual footer card in the right panel."""
    if completed and review_count > 0:
        st.markdown(
            f'<div class="tip-card review-tip">'
            f'<div class="tip-icon">⚠️</div>'
            f'<div>'
            f'<div class="tip-title">How to review {review_count} flagged item{"s" if review_count != 1 else ""}</div>'
            f'<p>Rows with missing or shared dimensions are highlighted in <b>red</b> in the Excel Master Log. '
            f'Open the workbook, check the embedded whiteboard photo in Column J, and fill in Column H.</p>'
            f'</div>'
            f'</div>',
            unsafe_allow_html=True,
        )
    elif completed:
        st.markdown(
            '<div class="tip-card ok-tip">'
            '<div class="tip-icon">✓</div>'
            '<div>'
            '<div class="tip-title">All extracted rows passed validation</div>'
            '<p>Every defect code and dimension passed automatic checks and was synced to SharePoint.</p>'
            '</div>'
            '</div>',
            unsafe_allow_html=True,
        )
    else:
        st.markdown(
            '<div class="tip-card">'
            '<div class="tip-icon">'
            '<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="#d97706" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">'
            '<path d="M12 3l1.9 5.8a2 2 0 0 0 1.3 1.3L21 12l-5.8 1.9a2 2 0 0 0-1.3 1.3L12 21l-1.9-5.8a2 2 0 0 0-1.3-1.3L3 12l5.8-1.9a2 2 0 0 0 1.3-1.3L12 3z"/>'
            '</svg>'
            '</div>'
            '<div>'
            '<div class="tip-title">Extraction tip</div>'
            '<p>For best results, use clear photos with the full whiteboard in frame.</p>'
            '</div>'
            '</div>',
            unsafe_allow_html=True,
        )