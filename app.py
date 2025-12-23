# Standard Imports
import streamlit as st

# Internal Imports
from src import run_excel_eda, func_eda_data_processing

# The setup
st.set_page_config(page_title="Data Preparation & EDA", layout="wide")

# The sidebars
st.sidebar.header("Navigation")
choice = st.sidebar.radio(
    "**Select a tool**", ["Data Preparation", "EDA Generation"], index=0, key="main_nav"
)

# The contents oh the
match choice:
    case "Data Preparation":
        func_eda_data_processing()
    case _:
        run_excel_eda()
