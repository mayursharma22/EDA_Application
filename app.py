# Standard Imports
import streamlit as st

# Internal Imports
from src import data_preperation, eda_generation

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
        data_preperation()
    case _:
        eda_generation()
