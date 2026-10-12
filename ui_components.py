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


def render_header():
    """The white header card: the Vonotec logo, a divider and the product name."""
    tagline = "Data Extraction and Master Log Automator"
    try:
        with open("vonotec.png", "rb") as f:
            logo_base64 = base64.b64encode(f.read()).decode()
        brand = f'<img src="data:image/png;base64,{logo_base64}" width="138" alt="VONOTEC">'
    except FileNotFoundError:
        brand = '<span class="brand-text">VONOTEC</span>'
    st.markdown(
        '<div class="header-container">'
        f'<div class="header-flex">{brand}<span class="product-name">Whiteboard AI</span></div>'
        f'<span class="header-tagline">{tagline}</span>'
        '</div>',
        unsafe_allow_html=True,
    )


def section_heading(eyebrow, title, badge=None):
    """A small orange label over a large title, with an optional status badge on the
    right. badge is (text, kind) where kind is "ready", "processing" or "complete"."""
    badge_html = ""
    if badge:
        text, kind = badge
        badge_html = f'<span class="status-badge {html.escape(kind)}"><span></span>{html.escape(text)}</span>'
    st.markdown(
        '<div class="section-heading"><div>'
        f'<p class="eyebrow">{html.escape(eyebrow)}</p>'
        f'<div class="section-title">{html.escape(title)}</div>'
        f'</div>{badge_html}</div>',
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
    with st.expander("How to use the Whiteboard Extractor", expanded=False):
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
