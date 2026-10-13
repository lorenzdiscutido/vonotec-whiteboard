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





# The "?" icon of the Help link in the header

_ICON_HELP = (

    '<svg aria-hidden="true" width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" '

    'stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="9"/>'

    '<path d="M9.5 9.2a2.6 2.6 0 0 1 5 .9c0 1.7-2.5 2.2-2.5 3.9M12 17h.01"/></svg>'

)



# The how-to text that opens from the Help link. One unbroken string with no indentation or

# blank lines, so Markdown keeps it as plain HTML.

_HELP_HTML = (

    '<div class="help-title">How to use the whiteboard extractor</div>'

    '<p>Add whiteboard photos from your device or from a SharePoint folder. The AI extracts the data '

    'and appends it to the Master Excel file.</p>'

    '<ol>'

    '<li>Choose where to import from: <b>Upload from device</b> or <b>Import from SharePoint</b>, then select your photos or folder.</li>'

    '<li>Choose the SharePoint folder to save into and give the batch a name. The name becomes the Excel file name.</li>'

    '<li>Click the <b>Process</b> button.</li>'

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

                '<div class="header-right">'

                '<span class="header-tagline">Data Extraction and Master Log Automator</span>'

                f'<details class="help-pop"><summary>{_ICON_HELP}Help</summary>'

                f'<div class="help-panel">{_HELP_HTML}</div></details>'

                '</div>',

                unsafe_allow_html=True,

            )

    return view





WORKFLOW_STEPS = ("Select photos", "Save destination", "Extract data")





def render_stepper(done, current):

    """A compact "1. Select photos > 2. Save destination > 3. Extract data" line.

    `done` is the set of finished step numbers (1 to 3) and `current` the step the person

    is on: finished steps are green with a check, the current one is blue, the rest are muted."""

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

    "complete", or a plain step label such as "Step 1 of 2". An optional

    description goes in a line under the title."""

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

    """The smaller heading of the source-specific content under the source switch

    (for example "Upload whiteboard photos")."""

    desc_html = f'<div class="source-subhead-desc">{html.escape(description)}</div>' if description else ""

    st.markdown(

        f'<div class="source-subhead"><div class="source-subhead-title">{html.escape(title)}</div>{desc_html}</div>',

        unsafe_allow_html=True,

    )





def card_heading(eyebrow, title, icon=_ICON_FILE, step=None):

    """The same label-over-title pair, smaller, for the top of a card. On the right: an

    optional step label such as "Step 2 of 2", then an optional small icon."""

    step_html = f'<span class="step-label">{html.escape(step)}</span>' if step else ""

    st.markdown(

        '<div class="card-heading"><div>'

        f'<p class="eyebrow">{html.escape(eyebrow)}</p>'

        f'<div class="card-title">{html.escape(title)}</div>'

        f'</div><div class="card-heading-right">{step_html}{icon}</div></div>',

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

                Add whiteboard photos from your device or from a SharePoint folder. The AI will automatically extract the data

                and append it to the Master Excel file.

                <ol>

                    <li>Choose where to import from: <b>Upload from device</b> or <b>Import from SharePoint</b>, then select your photos or folder.</li>

                    <li>Choose the SharePoint folder to save into and give the batch a name. The name becomes the Excel file name.</li>

                    <li>Click <b>Process photos</b>.</li>

                    <li>Follow the progress, messages and download on the right. The <b>Analytics Dashboard</b> picks up the new data by itself.</li>

                </ol>

            </div>

            """,

            unsafe_allow_html=True,

        ) 
