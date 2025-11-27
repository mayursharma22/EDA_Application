# Standard Imports
import runpy

# Third party imports
import streamlit as st

# Internal Imports
from src.eda_app.eda_excel_app import run_excel_eda


if __name__ == "__main__":
    st.set_page_config(page_title="Data Preparation & EDA", layout="wide")

    st.sidebar.header("Navigation")
    choice = st.sidebar.radio(
        "**Select a tool**",
        ["Data Preparation", "EDA Generation"],
        index=0,
        key="main_nav",
    )

    if choice == "Data Preparation":
        runpy.run_path("eda_data_Processing.py", run_name="__main__")

    else:
        run_excel_eda()
