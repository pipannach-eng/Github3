import base64
from pathlib import Path

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
      }
    </style>
    """,
    unsafe_allow_html=True,
)

base_dir = Path(__file__).parent


def read_text_file(file_name: str) -> str:
    return (base_dir / file_name).read_text(encoding="utf-8")


def safe_script(js_text: str) -> str:
    return js_text.replace("</script>", "<\\/script>")


def inline_asset(css_text: str, asset_path: str) -> str:
    file_path = base_dir / asset_path
    if not file_path.exists():
        return css_text

    encoded = base64.b64encode(file_path.read_bytes()).decode("ascii")
    data_uri = f"data:image/png;base64,{encoded}"
    return css_text.replace(f'url("./{asset_path}")', f"url('{data_uri}')")


required_files = ["index.html", "styles.css", "data.js", "app.js"]
missing_files = [name for name in required_files if not (base_dir / name).exists()]
if missing_files:
    st.error("ไม่พบไฟล์ที่จำเป็นสำหรับ dashboard")
    st.write(missing_files)
    st.stop()

try:
    html = read_text_file("index.html")
    css = read_text_file("styles.css")
    css = inline_asset(css, "assets/ecd-dashboard-animal-mascot.png")
    css = inline_asset(css, "assets/ecd-dashboard-mascot.png")
    data_js = safe_script(read_text_file("data.js"))
    app_js = safe_script(read_text_file("app.js"))

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
except Exception as exc:
    st.error("โหลดไฟล์ dashboard ไม่สำเร็จ")
    st.exception(exc)
    st.stop()

components.html(html, height=9000, scrolling=True)
