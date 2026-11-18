import streamlit as st

st.set_page_config(
    page_title="ClearView — Business Intelligence",
    page_icon="◈",
    layout="wide",
)

workspace = st.radio(
    "Choose a ClearView workspace",
    ["Finance", "Business operations"],
    horizontal=True,
)

if workspace == "Finance":
    from finance_workspace import render
else:
    from operations import render

render()
