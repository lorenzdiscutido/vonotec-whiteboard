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

    st.markdown(
        """
        <div class="instruction-text">
            Upload one or multiple whiteboard photos below. The AI will automatically extract the data 
            and append it to the Master Excel file. You can also drag and drop images directly into the uploader.
        </div>
        """, 
        unsafe_allow_html=True
    )
