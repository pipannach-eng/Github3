from pathlib import Path

import streamlit as st
import streamlit.components.v1 as components


st.set_page_config(
    page_title="G3 Dashboard",
    layout="wide",
)

base_dir = Path(__file__).parent

html = (base_dir / "index.html").read_text(encoding="utf-8")
css = (base_dir / "styles.css").read_text(encoding="utf-8")
data_js = (base_dir / "data.js").read_text(encoding="utf-8")
app_js = (base_dir / "app.js").read_text(encoding="utf-8")

html = html.replace(
    '<link rel="stylesheet" href="./styles.css">',
    f"<style>{css}</style>",
)
html = html.replace(
    '<script src="./data.js"></script>',
    f"<script>{data_js}</script>",
)
html = html.replace(
    '<script src="./app.js"></script>',
    f"<script>{app_js}</script>",
)

components.html(html, height=1900, scrolling=True)
