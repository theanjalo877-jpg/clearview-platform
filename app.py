import streamlit as st

st.set_page_config(
    page_title="ClearView Business Intelligence",
    page_icon="◈",
    layout="wide",
    initial_sidebar_state="collapsed",
)

def clear_finance_sample():
    st.session_state["finance_sample"] = False


workspace = st.radio(
    "Choose a ClearView workspace",
    ["Finance", "Business operations"],
    horizontal=True,
    on_change=clear_finance_sample,
)

if workspace == "Finance":
    from finance_workspace import render
else:
    from operations import render

render()
