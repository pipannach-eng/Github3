import streamlit as st
import streamlit.components.v1 as components


st.set_page_config(
    page_title="G3 Dashboard",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown(
    """
    <style>
      .block-container {
        max-width: 100% !important;
        padding: 0 !important;
      }

      header[data-testid="stHeader"] {
        background: rgba(255, 255, 255, 0.96) !important;
      }

      iframe {
        display: block;
        width: 100% !important;
        border: 0 !important;
      }
    </style>
    """,
    unsafe_allow_html=True,
)

# Load the dashboard as a real static page instead of inlining CSS/JS.
# This keeps colors, CSS backgrounds, images, and animations closest to local HTML.
dashboard_url = "https://raw.githack.com/pipannach-eng/Github3/streamlit-original-ui/index.html"

components.iframe(dashboard_url, height=9000, scrolling=True)
