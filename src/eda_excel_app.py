import os
import io
import tempfile
import hashlib
import pandas as pd
import streamlit as st
from src.eda_excel_generation import run as eda_excel_run

st.set_page_config(page_title="EDA Generation", layout="wide")

# ---------- helpers ----------
def _hash_file_like(file) -> str:
    data = file.getvalue() if hasattr(file, "getvalue") else file.read()
    if hasattr(file, "seek"):
        file.seek(0)
    return hashlib.sha256(data).hexdigest()

def _auto_detect_date_column(df: pd.DataFrame, sample_max: int = 5000, threshold: float = 0.6) -> int:
    """
    Return the index of the column most likely to be a date.
    Heuristics:
      - If a column is already datetime64, choose it immediately (first such column wins).
      - Else, try parsing up to 'sample_max' non-null rows; compute success ratio.
      - Pick the column with the highest ratio >= threshold.
      - If 'Date' exists, prefer it when ratios tie or no clear winner.
      - Fallback: index 0.
    """
    if df.empty or df.shape[1] == 0:
        return 0
    
    for i, c in enumerate(df.columns):
        if pd.api.types.is_datetime64_any_dtype(df[c]):
            return i

    best_idx = None
    best_score = -1.0
    for i, c in enumerate(df.columns):
        s = df[c].dropna()
        if s.empty:
            continue
        s = s.astype(str).head(sample_max)
        parsed = pd.to_datetime(s, errors="coerce", infer_datetime_format=True)
        score = parsed.notna().mean()
        if score > best_score:
            best_score = score
            best_idx = i

    date_name_idx = list(df.columns).index("Date") if "Date" in df.columns else None
    if best_idx is not None and best_score >= threshold:
        if date_name_idx is not None:
            s = df.iloc[:, date_name_idx].dropna().astype(str).head(sample_max)
            alt_score = pd.to_datetime(s, errors="coerce", infer_datetime_format=True).notna().mean() if not s.empty else 0
            if abs(alt_score - best_score) < 1e-9:
                return date_name_idx
        return best_idx

    if date_name_idx is not None:
        return date_name_idx
    return 0

@st.cache_data(show_spinner=False)
def build_eda_excel_bytes(file_bytes: bytes, file_name: str, params: dict) -> bytes:
    with tempfile.TemporaryDirectory() as tmpdir:
        src_path = os.path.join(tmpdir, file_name)
        with open(src_path, "wb") as f:
            f.write(file_bytes)

        out_path = os.path.join(tmpdir, "EDA_output.xlsx")
        gen_params = dict(params)
        gen_params["csv_path"] = src_path
        gen_params["output_path"] = out_path

        result_path = eda_excel_run(gen_params)

        with open(result_path, "rb") as fh:
            return fh.read()

# ---------- main ----------
def run_excel_eda():
    st.title("📈 EDA Generation")
    st.caption("Upload a CSV and configure parameters. Generates a formatted Excel workbook.")

    # Upload
    up = st.file_uploader("**Upload CSV**", type=["csv"], key="eda_uploader")
    if not up:
        st.info("Upload a CSV file to begin.")
        return

    df = pd.read_csv(up)
    st.markdown("##### Preview:")
    st.dataframe(df.head())
    st.markdown("##### Select Key Fields")
    auto_date_idx = _auto_detect_date_column(df)
    date_var = st.selectbox(
        "Date column",
        options=df.columns,
        index=auto_date_idx,
        key="eda_date_var"
    )

    grain_label = st.radio(
        "Date granularity",
        ["Daily","Weekly", "Monthly"],
        index=1,
        horizontal=True,
        key="eda_date_grain_label"
    )
    grain_map = {"Weekly": "weekly", "Monthly": "monthly", "Daily": "daily"}
    date_grain = grain_map[grain_label]

    ################# Newly Added Piece to detect Week Start Day ###################

    if grain_label == "Weekly":
        _dates = pd.to_datetime(df[date_var], errors="coerce").dropna()
        weekday_map = {
            0: "Monday", 1: "Tuesday", 2: "Wednesday",
            3: "Thursday", 4: "Friday", 5: "Saturday", 6: "Sunday",
        }
        reverse_weekday_map = {v: k for k, v in weekday_map.items()}

        selected_week_start = None

        if _dates.empty:
            st.caption("Could not detect a consistent week start day.")
            selected_week_start = st.selectbox(
                "Choose Week Start Day (will align/floor dates to this day)",
                options=list(reverse_weekday_map.keys()),
                index=0,
                key="eda_weekstart_select"
            )
            st.caption(f"The selected week start day is: **{selected_week_start}**")
        else:
            weekdays = _dates.dt.dayofweek.unique()
            if len(weekdays) == 1:
                detected = weekday_map[int(weekdays[0])]
                st.caption(f"In the dataset, Week Start Day is: **{detected}**")

                change_choice = st.radio(
                    "Do you want to change the Week Start Day?",
                    ["No", "Yes"],
                    index=0,
                    horizontal=True,
                    key="eda_change_weekstart"
                )
                if change_choice == "Yes":
                    selected_week_start = st.selectbox(
                        "Choose new Week Start Day",
                        options=list(reverse_weekday_map.keys()),
                        index=int(weekdays[0]),
                        key="eda_weekstart_select"
                    )
                    st.caption(f"The new week start day is: **{selected_week_start}**")
            else:
                st.caption("Could not detect a consistent week start day.")
                selected_week_start = st.selectbox(
                    "Choose Week Start Day (will align/floor dates to this day)",
                    options=list(reverse_weekday_map.keys()),
                    index=0,
                    key="eda_weekstart_select"
                )
                st.caption(f"The selected week start day is: **{selected_week_start}**")

     ################# Newly Added Piece to detect Week Start Day ###################

    default_metric_idx = list(df.columns).index("Metrics") if "Metrics" in df.columns else 0
    default_value_idx  = list(df.columns).index("Values")  if "Values"  in df.columns else 0

    metric_var = st.selectbox(
        "Metric column",
        options=df.columns,
        index=default_metric_idx,
        key="eda_metric_var"
    )
    value_var = st.selectbox(
        "Value column",
        options=df.columns,
        index=default_value_idx,
        key="eda_value_var"
    )

    metric_names = (
        df[metric_var].dropna().astype(str).str.strip().unique().tolist()
        if metric_var in df.columns else []
    )
    cost_metric_options = ["(None)"] + sorted(metric_names)
    cost_metric = st.selectbox(
        "Cost/Spend metric (optional)",
        options=cost_metric_options,
        index=0,
        key="eda_cost_metric"
    )
    cost_var = "" if cost_metric == "(None)" else cost_metric

    st.markdown("##### Select Breakdown Fields")

    dim_exclude = {date_var, metric_var, value_var}
    dim_candidates = [c for c in df.columns if c not in dim_exclude]

    qc_vars = st.multiselect(
        "Select column(s) to split data into individual Excel tab",
        options=dim_candidates,
        key="eda_qc_vars"
    )
    breakdown = st.multiselect(
        "Select Column(s) to split metrics by (e.g., Region, Subchannel)",
        options=[c for c in dim_candidates if c not in qc_vars],
        key="eda_breakdown"
    )

    params = {
        "date_var": date_var,
        "date_grain": date_grain,
        "QC_variables": qc_vars,
        "columns_breakdown": breakdown,
        "metrics": [],
        "metric_var": metric_var,
        "value_var": value_var,
        "cost_var": cost_var,
    }

    if selected_week_start:
        params["week_start_day"] = selected_week_start

       
    # ----- Color Customization -----
    st.markdown("#### Customize Colors")

    if "graph_colors" not in st.session_state:
        st.session_state.graph_colors = []
    if "tab_color" not in st.session_state:
        st.session_state.tab_color = "#12295D"

    # Default colors
    default_graph_colors = [
        '#12295D', '#00CACF', '#5B19C4', '#60608D',
        '#FFDC69', '#FF644C', '#06757E', '#996DDF', '#A2F9FB',
        '#41547D', '#FFAA00', '#2B49A6', '#439CA3', '#AEAEBC',
        '#E5E5E9', '#FFF3CD', '#111D23', '#F23A1D'
    ]
    default_tab_color = "#12295D"

    left, right = st.columns(2)

    # ----- Graph Colors -----
    with left:
        graph_choice = st.radio("**Choose Color for Visual Formatting**", ["Use Default Colors", "Pick Your Own"], key="graph_choice")

        if graph_choice == "Use Default Colors":
            st.session_state.graph_colors = default_graph_colors.copy()
            st.write("Using default graph colors:")
            selected_html = "<div style='display:flex;flex-wrap:wrap;'>"
            for color in st.session_state.graph_colors:
                selected_html += f"<div style='background:{color};width:40px;height:40px;margin:2px;border:1px solid #ccc;border-radius:4px;'></div>"
            selected_html += "</div>"
            st.markdown(selected_html, unsafe_allow_html=True)
        else:
            if st.session_state.graph_colors == default_graph_colors:
                st.session_state.graph_colors = []

            graph_color = st.color_picker("Pick Graph Color", value="#12295D")
            if st.button("Add Picked Color"):
                if graph_color not in st.session_state.graph_colors:
                    st.session_state.graph_colors.append(graph_color)

            def update_graph_colors():
                updated = [c.strip() for c in st.session_state.graph_input.split(",") if c.strip().startswith("#") and len(c.strip()) == 7]
                st.session_state.graph_colors = updated

            st.text_input("Insert HEX Colors (comma seperated)", value=",".join(st.session_state.graph_colors), key="graph_input", on_change=update_graph_colors)

            st.write("Selected Colors:")
            selected_html = "<div style='display:flex;flex-wrap:wrap;'>"
            for color in st.session_state.graph_colors:
                selected_html += f"<div style='background:{color};width:40px;height:40px;margin:2px;border:1px solid #ccc;border-radius:4px;'></div>"
            selected_html += "</div>"
            st.markdown(selected_html, unsafe_allow_html=True)

    # ----- Tab Color -----
    with right:
        tab_choice = st.radio("**Choose Color for Excel Formatting**", ["Use Default Color", "Pick Your Own"], key="tab_choice")

        if tab_choice == "Use Default Color":
            st.session_state.tab_color = default_tab_color
            st.write("Using default excel color:")
            st.markdown(f"<div style='background:{default_tab_color};width:40px;height:40px;margin:2px;border:1px solid #ccc;border-radius:4px;'></div>", unsafe_allow_html=True)
        else:
            if st.session_state.tab_color == default_tab_color:
                st.session_state.tab_color = "#12295D"

            tab_color_picker = st.color_picker("Pick Tab Color", value=st.session_state.tab_color)
            if st.button("Set Tab Color"):
                st.session_state.tab_color = tab_color_picker

            def update_tab_color():
                val = st.session_state.tab_input.strip()
                if val.startswith("#") and len(val) == 7:
                    st.session_state.tab_color = val

            st.text_input("Insert HEX Color", value=st.session_state.tab_color, key="tab_input", on_change=update_tab_color)

            st.write("Selected Color:")
            st.markdown(f"<div style='background:{st.session_state.tab_color};width:40px;height:40px;margin:2px;border:1px solid #ccc;border-radius:4px;'></div>", unsafe_allow_html=True)

    params["graph_colors"] = st.session_state.graph_colors
    params["tab_color"] = st.session_state.tab_color


    # Export
    st.markdown("#### Export")
    _ = _hash_file_like(up)

    with st.spinner("Preparing workbook..."):
        excel_bytes = build_eda_excel_bytes(
            file_bytes=up.getvalue(),
            file_name=up.name,
            params=params
        )

    ####### This is Temporary, since file downlaod is blocked #############
    st.markdown("##### Final Preview:")
    bio = io.BytesIO(excel_bytes)

    try:
        xls = pd.ExcelFile(bio, engine="openpyxl")
        sheet_names = xls.sheet_names
        sheet_to_show = st.selectbox("Select sheet to preview", sheet_names, index=0)
        df_preview = pd.read_excel(xls, sheet_name=sheet_to_show, engine="openpyxl")
        st.dataframe(df_preview)
        st.caption(
            f"Showing all {len(df_preview):,} rows of '{sheet_to_show}' "
            f"(columns: {df_preview.shape[1]})"
        )

    except Exception as e:
        st.error(f"Could not open the generated workbook for preview: {e}")
    ####### This is Temporary, since file downlaod is blocked #############

    st.download_button(
        label="🚀 Generate & Download EDA Workbook",
        data=excel_bytes,
        file_name="EDA_Final_Output.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="eda_generate_and_download"
    )

if __name__ == "__main__":
    run_excel_eda()