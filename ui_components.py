import base64
import streamlit as st

def load_local_css(file_name="style.css"):
    """Injects the external CSS file into the Streamlit app layout."""
    try:
        with open(file_name) as f:
            st.markdown(f'<style>{f.read()}</style>', unsafe_allow_html=True)
    except FileNotFoundError:
        pass

def render_header():
    """Renders the HTML layout containing the Vonotec logo and title strings."""
    try:
        with open("vonotec.png", "rb") as f:
            logo_base64 = base64.b64encode(f.read()).decode()
        st.markdown(
            f"""
            <div class="header-container">
                <div class="header-flex">
                    <img src="data:image/png;base64,{logo_base64}" width="180">
                    <div>
                        <h1 class="main-title">Whiteboard AI</h1>
                        <p class="sub-title">Data Extraction and Master Log Automator</p>
                    </div>
                </div>
            </div>
            """, 
            unsafe_allow_html=True
        )
    except FileNotFoundError:
        st.markdown(
            """
            <div class="header-container">
                <h1 class="main-title">VONOTEC Whiteboard AI</h1>
                <p class="sub-title">Data Extraction and Master Log Automator</p>
            </div>
            """, 
            unsafe_allow_html=True
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
