import streamlit as st
import runpy

st.set_page_config(page_title="Data Preparation & EDA", layout="wide")

st.sidebar.header("Navigation")
choice = st.sidebar.radio(
    "**Select a tool**",
    ["Data Preparation", 
     "EDA Generation"
     ],
    index=0,
    key="main_nav"
)

if choice == "Data Preparation":
    runpy.run_path("eda_data_Processing.py", run_name="__main__")

else:
    from eda_excel_app import run_excel_eda
    run_excel_eda()