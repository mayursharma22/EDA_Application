# Standard Imports
import streamlit as st

# Internal Imports
from src import data_preperation, eda_generation

# Setup
st.set_page_config(page_title="Data Preparation & EDA", layout="wide")

# Navigation helpers
TOOLS = ["Data Preparation", "EDA Generation"]


def _init_nav_state():
    if "effective_nav" not in st.session_state:
        st.session_state.effective_nav = TOOLS[0]
    if "pending_nav_target" not in st.session_state:
        st.session_state.pending_nav_target = None
    if "nav_confirm_open" not in st.session_state:
        st.session_state.nav_confirm_open = False


def _has_work_in_tool(tool_name: str) -> bool:
    ss = st.session_state

    if tool_name == "Data Preparation":
        return bool(
            ss.get("raw_data")
            or ss.get("preprocessed_data")
            or ss.get("grouped_data")
            or ss.get("schema_entries")
            or ss.get("rename_entries_df")
            or ss.get("rename_entries_melted")
            or (getattr(ss.get("final_melted"), "empty", True) is False)
            or (getattr(ss.get("final_grouped"), "empty", True) is False)
        )

    return bool(
        ss.get("eda_uploader") is not None
        or ss.get("eda_date_var")
        or ss.get("eda_metric_var")
        or ss.get("eda_value_var")
        or ss.get("eda_qc_vars")
        or ss.get("eda_breakdown")
        or ss.get("graph_colors")
        or ss.get("tab_color")
    )


def _reset_tool_state(prev_tool: str):
    preserve = {
        "effective_nav": st.session_state.get("effective_nav"),
        "pending_nav_target": st.session_state.get("pending_nav_target"),
        "nav_confirm_open": st.session_state.get("nav_confirm_open"),
    }
    st.session_state.clear()
    try:
        st.cache_data.clear()
    except Exception:
        pass
    for k, v in preserve.items():
        st.session_state[k] = v


# Dialog Box
@st.dialog("Confirm Action", width="medium")
def confirm_switch_dialog():
    st.markdown(
        "Switching between tools might lose your work."
        " **Are you sure you want to proceed?**",
        unsafe_allow_html=True,
    )
    c1, c2 = st.columns(2)
    with c1:
        if st.button("Proceed", type="primary"):
            prev = st.session_state.effective_nav
            target = st.session_state.pending_nav_target or prev

            _reset_tool_state(prev)
            st.session_state.effective_nav = target
            st.session_state.nav_confirm_open = False
            st.session_state.pending_nav_target = None
            st.session_state.main_nav = target
            st.rerun()

    with c2:
        if st.button("Cancel"):
            st.session_state.nav_confirm_open = False
            st.session_state.pending_nav_target = None
            st.session_state.main_nav = st.session_state.effective_nav
            st.rerun()


def _on_nav_change():
    """
    Radio change callback:
    - If user has worked in current tool: open dialog & do NOT switch (no st.rerun here).
    - If no work in current tool: switch immediately (no st.rerun needed in callback).
    """
    new_choice = st.session_state.get("main_nav", TOOLS[0])
    current = st.session_state.effective_nav
    if new_choice == current:
        return

    if _has_work_in_tool(current):
        st.session_state.pending_nav_target = new_choice
        st.session_state.nav_confirm_open = True
        st.session_state.main_nav = current
        confirm_switch_dialog()
    else:
        st.session_state.effective_nav = new_choice
        st.session_state.nav_confirm_open = False
        st.session_state.pending_nav_target = None
        st.session_state.main_nav = new_choice


# Navigation
st.sidebar.header("Navigation")
_init_nav_state()

if "main_nav" not in st.session_state:
    st.session_state["main_nav"] = st.session_state.effective_nav
else:
    if st.session_state.get("nav_confirm_open") is False:
        st.session_state["main_nav"] = st.session_state.effective_nav

st.sidebar.radio(
    "**Select a tool**",
    TOOLS,
    key="main_nav",
    on_change=_on_nav_change,
)


# Render the tool
if st.session_state.effective_nav == "Data Preparation":
    data_preperation()
else:
    eda_generation()
