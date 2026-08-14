from pathlib import Path

import streamlit as st
import streamlit.components.v1 as components


st.set_page_config(
    page_title="G3 Dashboard",
    layout="wide",
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


try:
    html = read_text_file("index.html")
    css = read_text_file("styles.css")
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

components.html(html, height=1900, scrolling=True)
