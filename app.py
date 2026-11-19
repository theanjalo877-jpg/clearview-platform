from pathlib import Path

import streamlit as st
import streamlit.components.v1 as components

st.set_page_config(
    page_title="ClearView Business Intelligence",
    page_icon="◈",
    layout="wide",
    initial_sidebar_state="collapsed",
)

def clear_finance_sample():
    st.session_state["finance_sample"] = False


masthead = st.container()

workspace = st.radio(
    "Choose a ClearView workspace",
    ["Finance", "Business operations"],
    horizontal=True,
    on_change=clear_finance_sample,
)

scene = "finance" if workspace == "Finance" else "operations"
masthead_html = Path(__file__).with_name("masthead.html").read_text(encoding="utf-8")
with masthead:
    components.html(masthead_html.replace("__SCENE__", scene), height=290)

if workspace == "Finance":
    from finance_workspace import render
else:
    from operations import render

render()
