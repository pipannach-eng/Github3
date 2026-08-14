import base64
from pathlib import Path

import streamlit as st
import streamlit.components.v1 as components


st.set_page_config(
    page_title="G3 Dashboard",
    layout="wide",
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

required_files = ["index.html", "styles.css", "data.js", "app.js"]
missing_files = [name for name in required_files if not (base_dir / name).exists()]

if missing_files:
    st.error("ไม่พบไฟล์ที่จำเป็นสำหรับ dashboard")
    st.write(missing_files)
    st.stop()


def read_text_file(file_name: str) -> str:
    return (base_dir / file_name).read_text(encoding="utf-8")


def safe_script(js_text: str) -> str:
    # Prevent embedded JS text from accidentally closing the surrounding script tag.
    return js_text.replace("</script>", "<\\/script>")


def inline_asset(css_text: str, asset_path: str) -> str:
    file_path = base_dir / asset_path
    if not file_path.exists():
        return css_text

    encoded = base64.b64encode(file_path.read_bytes()).decode("ascii")
    data_uri = f"data:image/png;base64,{encoded}"
    return css_text.replace(f'url("./{asset_path}")', f"url('{data_uri}')")


STREAMLIT_DASHBOARD_FIX_CSS = """
<style>
  html,
  body {
    width: 100%;
    min-width: 1440px;
    overflow-x: auto;
    background:
      radial-gradient(circle at top left, rgba(255, 180, 99, 0.2), transparent 24%),
      radial-gradient(circle at right top, rgba(31, 183, 189, 0.18), transparent 20%),
      linear-gradient(180deg, #f8fdfe 0%, #eaf7f6 100%) !important;
  }

  .dashboard-layout {
    grid-template-columns: 260px minmax(0, 1fr) !important;
    gap: 24px !important;
    padding: 18px !important;
    min-height: 100vh !important;
  }

  .sidebar {
    display: flex !important;
    position: sticky !important;
    top: 18px !important;
    height: calc(100vh - 36px) !important;
    background: linear-gradient(180deg, #0f7b7f 0%, #07676f 100%) !important;
    color: #ffffff !important;
    box-shadow: 0 24px 60px rgba(42, 106, 116, 0.16) !important;
    opacity: 1 !important;
  }

  .sidebar *,
  .nav-item,
  .brand-copy h2,
  .sidebar-card strong {
    color: #ffffff !important;
    opacity: 1 !important;
  }

  .brand-label,
  .sidebar-card span,
  .sidebar-card small,
  .sidebar-mini-chart p {
    color: rgba(235, 255, 255, 0.82) !important;
  }

  .nav-item.active {
    background: rgba(255, 255, 255, 0.17) !important;
  }

  .hero,
  .filters-panel,
  .secondary-kpi,
  .policy-action-card.is-hidden {
    display: none !important;
  }

  .page-shell {
    padding: 0 6px 40px 0 !important;
  }

  .executive-board {
    display: block !important;
    margin-top: 0 !important;
  }
</style>
"""


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
    html = html.replace("</head>", f"{STREAMLIT_DASHBOARD_FIX_CSS}</head>")
except Exception as exc:
    st.error("โหลดไฟล์ dashboard ไม่สำเร็จ")
    st.exception(exc)
    st.stop()

components.html(html, height=1900, scrolling=True)
