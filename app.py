from pathlib import Path

import streamlit as st
import streamlit.components.v1 as components

st.set_page_config(
    page_title="ClearView Business Intelligence",
    page_icon="◈",
    layout="wide",
    initial_sidebar_state="collapsed",
)

def clear_samples():
    st.session_state["finance_sample"] = False
    st.session_state["operations_sample"] = False


masthead = st.container()

workspace = st.radio(
    "Choose a ClearView workspace",
    ["Finance", "Business operations"],
    horizontal=True,
    on_change=clear_samples,
)

scene = "finance" if workspace == "Finance" else "operations"
masthead_html = Path(__file__).with_name("masthead.html").read_text(encoding="utf-8")
with masthead:
    components.html(masthead_html.replace("__SCENE__", scene), height=290)

if scene == "finance":
    demo_name = "ClearView_Finance_Demo.xlsx"
    demo_mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
else:
    demo_name = "ClearView_Business_Operations_Demo.docx"
    demo_mime = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"

st.download_button(
    "Download client demo",
    data=(Path(__file__).with_name("data") / demo_name).read_bytes(),
    file_name=demo_name,
    mime=demo_mime,
    on_click="ignore",
    help="Fictional Cedarline Fixtures data. Upload the downloaded file in this workspace to try the client presentation.",
)

if workspace == "Finance":
    from finance_workspace import render
else:
    from operations import render

render()
