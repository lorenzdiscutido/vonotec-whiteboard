import base64
import html

import streamlit as st

# Inline icons (the same paths the design uses)
_ICON_LAYERS = (
    '<svg width="23" height="23" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" '
    'stroke-linecap="round" stroke-linejoin="round"><path d="m12 3 9 5-9 5-9-5 9-5Zm-7.5 9 7.5 4 7.5-4M4.5 16l7.5 4 7.5-4"/></svg>'
)
_ICON_SPARK = (
    '<svg width="18" height="18" viewBox="0 0 24 24" fill="currentColor"><path d="m12 3 1.2 4.1L17 9l-3.8 1.9L12 15l-1.2-4.1L7 9l3.8-1.9L12 3ZM5 15l.7 2.3L8 18.5l-2.3 1.2L5 22l-.7-2.3L2 18.5l2.3-1.2L5 15Zm14-3 .7 2.3 2.3 1.2-2.3 1.2L19 19l-.7-2.3-2.3-1.2 2.3-1.2L19 12Z"/></svg>'
)


def load_local_css(file_name="style.css"):
    """Injects the external CSS file into the Streamlit app layout."""
    try:
        with open(file_name, encoding="utf-8") as f:
            st.markdown(f'<style>{f.read()}</style>', unsafe_allow_html=True)
    except FileNotFoundError:
        pass


def _brand_html():
    """The logo, a thin divider and the product name (left of the header bar)."""
    try:
        with open("vonotec.png", "rb") as f:
            logo_base64 = base64.b64encode(f.read()).decode()
        logo = f'<img src="data:image/png;base64,{logo_base64}" width="138" alt="VONOTEC">'
    except FileNotFoundError:
        logo = '<span class="brand-text">VONOTEC</span>'
    return f'<div class="brand">{logo}<span class="product-name">Whiteboard AI</span></div>'


def render_header(views):
    """One white header bar with three parts: the brand on the left, the view switch
    (a radio, drawn as two pills) in the middle and the tagline on the right.
    `views` is the list of view names; returns the one that is chosen."""
    with st.container(key="app_header"):
        try:
            col_brand, col_nav, col_tagline = st.columns([1.2, 1, 1.2], vertical_alignment="center")
        except TypeError:  # older Streamlit without vertical_alignment
            col_brand, col_nav, col_tagline = st.columns([1.2, 1, 1.2])

        with col_brand:
            st.markdown(_brand_html(), unsafe_allow_html=True)
        with col_nav:
            with st.container(key="view_switch"):
                view = st.radio(
                    "View", views,
                    horizontal=True, label_visibility="collapsed", key="active_view",
                )
        with col_tagline:
            st.markdown(
                '<div class="header-tagline">Data Extraction and Master Log Automator</div>',
                unsafe_allow_html=True,
            )
    return view


def section_heading(eyebrow, title, badge=None, step=None):
    """A small orange label over a large title. On the right there is either a status
    badge, badge = (text, kind) with kind "ready", "processing" or "complete", or a plain
    step label such as "Step 1 of 2"."""
    right = ""
    if badge:
        text, kind = badge
        right = f'<span class="status-badge {html.escape(kind)}"><span></span>{html.escape(text)}</span>'
    elif step:
        right = f'<span class="step-label">{html.escape(step)}</span>'
    st.markdown(
        '<div class="section-heading"><div>'
        f'<p class="eyebrow">{html.escape(eyebrow)}</p>'
        f'<h2>{html.escape(title)}</h2>'
        f'</div>{right}</div>',
        unsafe_allow_html=True,
    )


def card_heading(eyebrow, title):
    """The same label-over-title pair, smaller, for the top of a card."""
    st.markdown(
        f'<p class="eyebrow">{html.escape(eyebrow)}</p>'
        f'<div class="card-title">{html.escape(title)}</div>',
        unsafe_allow_html=True,
    )


def render_empty_results():
    """Shown in the Batch results column until the first batch is run."""
    st.markdown(
        '<div class="sync-card">'
        f'<span class="sync-icon">{_ICON_LAYERS}</span>'
        '<div><strong>Ready to process a new batch</strong>'
        '<p>Messages, downloads and sync status appear here once you extract a batch.</p></div>'
        '</div>',
        unsafe_allow_html=True,
    )


def render_tip():
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
                Upload one or multiple whiteboard photos. The AI will automatically extract the data
                and append it to the Master Excel file. You can also drag and drop images directly into the uploader.
                <ol>
                    <li>Add your whiteboard photos (a preview of each appears below the uploader).</li>
                    <li>Choose the SharePoint folder and give the batch a name. The name becomes the Excel file name.</li>
                    <li>Click <b>Extract Data &amp; Update Excel</b>.</li>
                    <li>Follow the progress, messages and download on the right. The <b>Analytics Dashboard</b> picks up the new data by itself.</li>
                </ol>
            </div>
            """,
            unsafe_allow_html=True
        )
