import base64
import html
from pathlib import Path
import streamlit as st

# Inline icons (the same paths the design uses)
_ICON_LAYERS = (
    '<svg width="23" height="23" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" '
    'stroke-linecap="round" stroke-linejoin="round"><path d="m12 3 9 5-9 5-9-5 9-5Zm-7.5 9 7.5 4 7.5-4M4.5 16l7.5 4 7.5-4"/></svg>'
)
_ICON_SPARK = (
    '<svg width="18" height="18" viewBox="0 0 24 24" fill="currentColor"><path d="m12 3 1.2 4.1L17 9l-3.8 1.9L12 15l-1.2-4.1L7 9l3.8-1.9L12 3ZM5 15l.7 2.3L8 18.5l-2.3 1.2L5 22l-.7-2.3L2 18.5l2.3-1.2L5 15Zm14-3 .7 2.3 2.3 1.2-2.3 1.2L19 19l-.7-2.3-2.3-1.2 2.3-1.2L19 12Z"/></svg>'
)

# Gray file icon on the right of the "Save settings" card heading
_ICON_FILE = (
    '<svg aria-hidden="true" width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="#9ba8b9" '
    'stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round">'
    '<path d="M7 3.5h6l4 4V20H7V3.5Zm6 0v4h4M9.5 12h5m-5 3h5"/></svg>'
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
    """One white header card with three equal columns: the brand on the left,
    the view switch (a radio, drawn as two pills) in the middle and the
    tagline on the right. views is the list of view names; returns the one
    that is chosen."""
    with st.container(key="top_app_header"):
        try:
            col_left, col_center, col_right = st.columns([1, 1, 1], vertical_alignment="center")
        except TypeError:  # older Streamlit without vertical_alignment
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
                '<div class="header-tagline" style="text-align: right; color: #64748b; font-size: 12px; '
                'font-weight: 500; white-space: nowrap;">Data Extraction and Master Log Automator</div>',
                unsafe_allow_html=True,
            )
    return view


def section_heading(eyebrow, title, badge=None, step=None):
    """A small orange label over a large title. On the right there is either a
    status badge, badge = (text, kind) with kind "ready", "processing" or
    "complete", or a plain step label such as "Step 1 of 2"."""
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


def card_heading(eyebrow, title, icon=_ICON_FILE):
    """The same label-over-title pair, smaller, for the top of a card, with an
    optional small icon on the right."""
    st.markdown(
        '<div class="card-heading"><div>'
        f'<p class="eyebrow">{html.escape(eyebrow)}</p>'
        f'<div class="card-title">{html.escape(title)}</div>'
        f'</div>{icon}</div>',
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
                    <li>Click <b>Extract Data & Update Excel</b>.</li>
                    <li>Follow the progress, messages and download on the right. The <b>Analytics Dashboard</b> picks up the new data by itself.</li>
                </ol>
            </div>
            """,
            unsafe_allow_html=True,
        )
