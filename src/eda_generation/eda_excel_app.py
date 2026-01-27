# Standard Library
import os
import tempfile
import hashlib

# Third Party imports
import pandas as pd
import streamlit as st

# Internal imports
from .eda_excel_generation import run as eda_excel_run
from .eda_ppt_generation import run as eda_ppt_run

st.set_page_config(page_title="EDA Generation", layout="wide")

# ---------- helpers ----------
@st.cache_data(show_spinner=False)
def build_eda_ppt_bytes(file_bytes: bytes, file_name: str, params: dict, template_bytes: bytes | None) -> bytes:
    import os, tempfile
    with tempfile.TemporaryDirectory() as tmpdir:
        src_path = os.path.join(tmpdir, file_name)
        with open(src_path, "wb") as f:
            f.write(file_bytes)
        template_path = None
        if template_bytes:
            template_path = os.path.join(tmpdir, "template.pptx")
            with open(template_path, "wb") as tf:
                tf.write(template_bytes)
        df = pd.read_csv(src_path)
        out_path = os.path.join(tmpdir, "EDA_Deck.pptx")
        eda_ppt_run(params=params, template_path=template_path, df=df, output_path=out_path)
        with open(out_path, "rb") as fh:
            return fh.read()

def _get_sample_csv_bytes() -> bytes:
    sample_csv = (
        "Date,Brand,Indication,Subchannel,Segment,Publisher,Channel,Metrics,Values\n"
        "1/6/2025,Brand_A,Indication_A,Brand,DTC,Bing,Search,Clicks,321\n"
        "1/6/2025,Brand_A,Indication_A,Brand,DTC,Bing,Search,Cost,1823.25\n"
        "1/6/2025,Brand_A,Indication_A,Brand,DTC,Bing,Search,Impressions,7164\n"
        "1/6/2025,Brand_A,Indication_A,Brand,DTC,Google,Search,Clicks,1695\n"
        "1/6/2025,Brand_A,Indication_A,Brand,DTC,Google,Search,Cost,12557.04\n"
        "1/6/2025,Brand_B,Indication_A,Banner,HCP,Meta,Social,Clicks,3749\n"
        "1/6/2025,Brand_B,Indication_A,Banner,HCP,Meta,Social,Cost,12357.00045\n"
        "1/6/2025,Brand_B,Indication_A,Banner,HCP,Meta,Social,Impressions,1895173\n"
        "1/13/2025,Brand_A,Indication_A,Banner,HCP,Meta,Social,Clicks,2256\n"
        "1/13/2025,Brand_A,Indication_A,Banner,HCP,Meta,Social,Cost,12458.43036\n"
    )
    return sample_csv.encode("utf-8")


def _hash_file_like(file) -> str:
    data = file.getvalue() if hasattr(file, "getvalue") else file.read()
    if hasattr(file, "seek"):
        file.seek(0)
    return hashlib.sha256(data).hexdigest()

def _auto_detect_date_column(
    df: pd.DataFrame, sample_max: int = 5000, threshold: float = 0.6
) -> int:
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
            alt_score = (
                pd.to_datetime(s, errors="coerce", infer_datetime_format=True)
                .notna()
                .mean()
                if not s.empty
                else 0
            )
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


@st.dialog("Long Format Required", width="medium")
def long_format_required_dialog():
    st.markdown(
        """
**This file does not look like a *long format* dataset.**

A long-format CSV typically has:
- A date/time column (e.g., `Date`)
- One or more categorical dimensions (e.g., `Channel`, `SubChannel`, `Metrics`)
- Exactly **one** numeric **Values** column (e.g., `Values`)

Please upload a long-format CSV that has exactly one metric-like numeric column.
        """,
        unsafe_allow_html=True,
    )
    if st.button("Okay", type="primary"):
        st.session_state["uploader_nonce"] = (
            st.session_state.get("uploader_nonce", 0) + 1
        )
        st.rerun()


def _is_monotonic_like(s: pd.Series) -> bool:
    """
    Treat as categorical/ID-like if the numeric series is monotonic non-decreasing
    (e.g., 1,2,3,... or 500,501,...) after sorting by natural keys.
    """
    if not pd.api.types.is_numeric_dtype(s):
        return False
    s = s.dropna()
    if s.empty:
        return False
    return bool(s.is_monotonic_increasing or s.is_monotonic_decreasing)


def _split_numeric_id_vs_metric(
    df: pd.DataFrame,
    sort_keys: list,
    monotonic_checks: bool = True,
) -> tuple[list[str], list[str]]:
    """
    Classify numeric columns using ONLY monotonicity (no ratio/density):
      - id_like    : integer dtype AND monotonic after sorting by 'sort_keys'
      - metric_like: everything else
    """
    id_like, metric_like = [], []

    if sort_keys:
        mask = pd.Series(True, index=df.index)
        for k in sort_keys:
            if k in df.columns:
                mask &= df[k].notna()
        sorted_df = df.loc[mask].sort_values(sort_keys, kind="mergesort")
    else:
        sorted_df = df

    for col in df.select_dtypes(include="number").columns:
        s_sorted = sorted_df[col].dropna()
        if (
            monotonic_checks
            and pd.api.types.is_integer_dtype(s_sorted)
            and _is_monotonic_like(s_sorted)
        ):
            id_like.append(col)
        else:
            metric_like.append(col)

    return id_like, metric_like


# ---------- main ----------
def eda_generation():
    st.title("📈 EDA Generation")
    st.caption(
        "Upload a CSV and configure parameters. Generates a formatted Excel workbook."
    )

    # ------------------------------- Upload File ---------------------------
    up = st.file_uploader(
        "**Upload CSV**",
        type=["csv"],
        key=f"eda_uploader_{st.session_state.get('uploader_nonce', 0)}",
    )

    if not up:
        st.info("Upload a CSV file to begin")

        with st.expander("📄 **Sample file format** ", expanded=True):

            import io
            sample_bytes = _get_sample_csv_bytes()
            sample_df = pd.read_csv(io.BytesIO(sample_bytes))
            st.dataframe(sample_df, use_container_width=True)

        return

    if "_use_sample_csv_bytes" in st.session_state and st.session_state["_use_sample_csv_bytes"]:
        import io
        up = io.BytesIO(st.session_state["_use_sample_csv_bytes"])
        up.name = "eda_sample.csv"


    df = pd.read_csv(up)
    df.columns = [str(c).strip() for c in df.columns]
    is_unnamed = df.columns.str.match(r"(?i)^\s*unnamed")
    is_blank = df.columns.str.strip() == ""
    mask_unnamed_or_blank = is_unnamed | is_blank

    if mask_unnamed_or_blank.any():
        df = df.loc[:, ~mask_unnamed_or_blank]

    if mask_unnamed_or_blank.any():
        df = df.loc[:, ~mask_unnamed_or_blank]

    sort_keys = []
    try:
        auto_idx = _auto_detect_date_column(df)
        if 0 <= auto_idx < len(df.columns):
            sort_keys = [df.columns[auto_idx]]
    except Exception:
        date_like = [c for c in df.columns if "date" in str(c).lower()]
        if date_like:
            sort_keys = [date_like[0]]

    id_like, metric_like = _split_numeric_id_vs_metric(
        df=df,
        sort_keys=sort_keys,
        monotonic_checks=True,
    )

    metric_like_count = len(metric_like)

    if metric_like_count != 1:
        long_format_required_dialog()
        return

    for c in id_like:
        df[c] = df[c].astype("category")

    st.markdown("#### Preview:")
    st.dataframe(df.head())

    # ------------------------------- Upload File ---------------------------

    # ----------------------------- Date Selection --------------------------
    st.markdown("#### Select Key Fields")
    auto_date_idx = _auto_detect_date_column(df)
    date_var = st.selectbox(
        "Date column", options=df.columns, index=auto_date_idx, key="eda_date_var"
    )

    grain_label = st.radio(
        "Date granularity",
        ["Daily", "Weekly", "Monthly"],
        index=1,
        horizontal=True,
        key="eda_date_grain_label",
    )
    grain_map = {"Weekly": "weekly", "Monthly": "monthly", "Daily": "daily"}
    date_grain = grain_map[grain_label]

    if grain_label == "Weekly":
        _dates = pd.to_datetime(df[date_var], errors="coerce").dropna()
        weekday_map = {
            0: "Monday",
            1: "Tuesday",
            2: "Wednesday",
            3: "Thursday",
            4: "Friday",
            5: "Saturday",
            6: "Sunday",
        }
        reverse_weekday_map = {v: k for k, v in weekday_map.items()}

        selected_week_start = None

        if _dates.empty:
            st.caption("Could not detect a consistent week start day.")
            selected_week_start = st.selectbox(
                "Choose Week Start Day (will align/floor dates to this day)",
                options=list(reverse_weekday_map.keys()),
                index=0,
                key="eda_weekstart_select",
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
                    key="eda_change_weekstart",
                )
                if change_choice == "Yes":
                    selected_week_start = st.selectbox(
                        "Choose new Week Start Day",
                        options=list(reverse_weekday_map.keys()),
                        index=int(weekdays[0]),
                        key="eda_weekstart_select",
                    )
                    st.caption(f"The new week start day is: **{selected_week_start}**")
            else:
                st.caption("Could not detect a consistent week start day.")
                selected_week_start = st.selectbox(
                    "Choose Week Start Day (will align/floor dates to this day)",
                    options=list(reverse_weekday_map.keys()),
                    index=0,
                    key="eda_weekstart_select",
                )
                st.caption(f"The selected week start day is: **{selected_week_start}**")

    # ----------------------------- Date Selection --------------------------

    # ------------------- Metric/Value and Cost Selection  -----------------

    def _find_index_case_insensitive(cols, candidates):
        cmap = [c.strip().lower() for c in cols]
        for cand in candidates:
            if cand.lower() in cmap:
                return cmap.index(cand.lower())
        return -1

    metric_idx = _find_index_case_insensitive(df.columns, ["metrics", "metric"])
    value_idx = _find_index_case_insensitive(df.columns, ["values", "value"])

    metric_options = ["(None)"] + list(df.columns)
    value_options = ["(None)"] + list(df.columns)

    metric_index = 0 if metric_idx == -1 else metric_idx + 1
    value_index = 0 if value_idx == -1 else value_idx + 1

    metric_var_sel = st.selectbox(
        "Metric column",
        options=metric_options,
        index=metric_index,
        key="eda_metric_var",
    )
    value_var_sel = st.selectbox(
        "Value column",
        options=value_options,
        index=value_index,
        key="eda_value_var",
    )

    is_metric_none = metric_var_sel == "(None)"
    is_value_none = value_var_sel == "(None)"
    same_col = (
        not is_metric_none and not is_value_none and metric_var_sel == value_var_sel
    )

    if same_col:
        metric_names = [metric_var_sel]
    elif not is_metric_none and (metric_var_sel in df.columns):
        metric_names = (
            df[metric_var_sel].dropna().astype(str).str.strip().unique().tolist()
        )
    else:
        metric_names = []

    cost_metric_options = ["(None)"] + sorted(metric_names)
    cost_metric = st.selectbox(
        "Cost/Spend metric (optional)",
        options=cost_metric_options,
        index=0,
        key="eda_cost_metric",
    )
    cost_var = "" if cost_metric == "(None)" else cost_metric

    st.markdown("#### Select Breakdown Fields")

    # dim_exclude = {date_var, metric_var, value_var}
    dim_exclude = {date_var}
    if not is_metric_none:
        dim_exclude.add(metric_var_sel)
    if not is_value_none:
        dim_exclude.add(value_var_sel)

    dim_candidates = [c for c in df.columns if c not in dim_exclude]

    # ------------------- Metric/Value and Cost Selection  -----------------

    # -------------------------- Breakdown Selection  ----------------------

    qc_vars = st.multiselect(
        "Select column(s) to split data into individual Excel tab",
        options=dim_candidates,
        key="eda_qc_vars",
    )
    breakdown = st.multiselect(
        "Select Column(s) to split metrics by (e.g., Region, Subchannel)",
        options=[c for c in dim_candidates if c not in qc_vars],
        key="eda_breakdown",
    )

    # -------------------- Breakdown Selection  ------------------#

    # ------------------------ Param ---------------------------- #

    params = {
        "date_var": date_var,
        "date_grain": date_grain,
        "QC_variables": qc_vars,
        "columns_breakdown": breakdown,
        "metrics": [],
        "metric_var": "",
        "value_var": "",
        "cost_var": cost_var,
    }
    if selected_week_start:
        params["week_start_day"] = selected_week_start

    if same_col:
        params["metrics"] = [metric_var_sel]
    else:
        if not is_metric_none:
            params["metric_var"] = metric_var_sel
        if not is_value_none:
            params["value_var"] = value_var_sel

    # ------------------------ Param ---------------------------- #

    # ------------------- Color Customization Section ----------- #
    st.markdown("#### Customize Colors")

    if "graph_colors" not in st.session_state:
        st.session_state.graph_colors = []
    if "tab_color" not in st.session_state:
        st.session_state.tab_color = "#12295D"

    # Default colors
    default_graph_colors = [
        "#12295D",
        "#00CACF",
        "#5B19C4",
        "#60608D",
        "#FFDC69",
        "#FF644C",
        "#06757E",
        "#996DDF",
        "#A2F9FB",
        "#41547D",
        "#FFAA00",
        "#2B49A6",
        "#439CA3",
        "#AEAEBC",
        "#E5E5E9",
        "#FFF3CD",
        "#111D23",
        "#F23A1D",
    ]
    default_tab_color = "#12295D"

    left, right = st.columns(2)

    # ---------------------- Graph Colors -----------------------
    #
    with left:

        def _reset_graph_warnings():
            st.session_state["_graph_invalid"] = []
            st.session_state["_graph_trimmed"] = []
            st.session_state["_graph_fixed_commas"] = False

        if "_graph_invalid" not in st.session_state:
            st.session_state["_graph_invalid"] = []
        if "_graph_trimmed" not in st.session_state:
            st.session_state["_graph_trimmed"] = []
        if "_graph_fixed_commas" not in st.session_state:
            st.session_state["_graph_fixed_commas"] = False
        if "_graph_choice_prev" not in st.session_state:
            st.session_state["_graph_choice_prev"] = None

        st.markdown(
            "<div style='font-size:18px; font-weight:600; margin-bottom:4px;'>Choose Color for Visual Formatting</div>",
            unsafe_allow_html=True,
        )
        graph_choice = st.radio(
            "",
            ["Use Default Colors", "Pick Your Own"],
            key="graph_choice",
            label_visibility="collapsed",
        )

        prev_choice = st.session_state["_graph_choice_prev"]
        if prev_choice is None or prev_choice != graph_choice:
            if graph_choice == "Use Default Colors":
                st.session_state.graph_colors = default_graph_colors.copy()
                st.session_state.graph_input = ""
            else:
                st.session_state.graph_input = ""
            _reset_graph_warnings()
            st.session_state["_graph_choice_prev"] = graph_choice

        if graph_choice == "Use Default Colors":
            st.session_state.graph_colors = default_graph_colors.copy()
            st.write("Using default graph colors:")
            selected_html = "<div style='display:flex;flex-wrap:wrap;'>"
            for color in st.session_state.graph_colors:
                selected_html += (
                    f"<div style='background:{color};width:40px;height:40px;"
                    "margin:2px;border:1px solid #ccc;border-radius:4px;'></div>"
                )
            selected_html += "</div>"
            st.markdown(selected_html, unsafe_allow_html=True)

        else:
            import re

            def _extract_hex_list(s):
                return [m.upper() for m in re.findall(r"\#[0-9A-Fa-f]{6}", s or "")]

            def _analyze_tokens(raw):
                pieces = [
                    p.strip() for p in re.split(r"[,\\s]+", raw or "") if p.strip()
                ]
                invalid, trimmed_hexes, fixed_commas = [], [], False
                for p in pieces:
                    matches = re.findall(r"\#[0-9A-Fa-f]{6}", p)
                    if not matches:
                        invalid.append(p)
                    else:
                        joined = "".join(matches)
                        if len(matches) >= 2 and joined == p:
                            fixed_commas = True
                        elif len(matches) == 1 and p == matches[0]:
                            pass
                        else:
                            for h in matches:
                                trimmed_hexes.append(h.upper())
                seen, deduped = set(), []
                for h in trimmed_hexes:
                    if h not in seen:
                        seen.add(h)
                        deduped.append(h)
                return invalid, deduped, fixed_commas

            def _dedup(seq):
                seen, out = set(), []
                for x in seq:
                    if x not in seen:
                        seen.add(x)
                        out.append(x)
                return out

            if "graph_colors" not in st.session_state:
                st.session_state.graph_colors = []
            if "graph_input" not in st.session_state:
                st.session_state.graph_input = ""
            if st.session_state.graph_colors == default_graph_colors:
                st.session_state.graph_colors = []
                st.session_state.graph_input = ""

            def update_graph_colors():
                raw = (st.session_state.graph_input or "").strip()
                if raw == "":
                    _reset_graph_warnings()
                    st.session_state.graph_input = ""
                    return
                hexes = _dedup(_extract_hex_list(raw))
                invalid, trimmed_hexes, fixed_commas = _analyze_tokens(raw)
                st.session_state.graph_colors = hexes
                st.session_state.graph_input = ",".join(hexes)
                st.session_state._graph_invalid = invalid
                st.session_state._graph_trimmed = trimmed_hexes
                st.session_state._graph_fixed_commas = fixed_commas

            # ---------- ROW 1: swatch (left) + Add Picked Color (right) on the SAME row ----------
            st.markdown(
                "<div style='font-size:15px; font-weight:600; margin-top:4px;'>Pick Graph Color</div>",
                unsafe_allow_html=True,
            )

            col_sw, col_add = st.columns([6, 4])
            with col_sw:
                graph_color = st.color_picker(
                    label="",
                    value="#12295D",
                    key="graph_color_picker",
                    label_visibility="collapsed",
                )
            with col_add:
                if st.button(
                    "Add Picked Color",
                    key="add_picked_color",
                    type="primary",
                    use_container_width=True,
                ):
                    c = (graph_color or "").strip().upper()
                    if re.fullmatch(r"\#[0-9A-Fa-f]{6}", c):
                        if c not in st.session_state.graph_colors:
                            st.session_state.graph_colors.append(c)
                            st.session_state.graph_input = ",".join(
                                st.session_state.graph_colors
                            )
                            _reset_graph_warnings()
                    else:
                        st.warning(
                            "Please enter a valid HEX color in the form #RRGGBB (e.g., #12295D)."
                        )

            # -------- ROW 2: Clear All (right) BEFORE text input (left) --------
            gi_col_input, gi_col_clear = st.columns([7, 2])

            with gi_col_input:
                st.text_input(
                    "Insert HEX Colors (comma separated)",
                    key="graph_input",
                    on_change=update_graph_colors,
                    placeholder="#12295D, #00CACF, #5B19C4",
                    help="Enter 7-character HEX colors like #AABBCC. Separate multiple colors with commas.",
                )

            def _clear_graph_section():
                st.session_state.graph_colors = []
                st.session_state.graph_input = ""
                _reset_graph_warnings()

            with gi_col_clear:
                st.markdown("<div style='height:28px'></div>", unsafe_allow_html=True)
                st.button(
                    "Clear All",
                    key="clear_graph_colors",
                    type="primary",
                    use_container_width=True,
                    on_click=_clear_graph_section,
                )

            # -------- messages & preview --------
            if st.session_state.get("_graph_fixed_commas"):
                st.warning(
                    "Missing commas detected between color codes. Your list was auto-corrected."
                )
            trimmed_list = st.session_state.get("_graph_trimmed", [])
            if trimmed_list:
                st.warning(
                    f"Extra characters were removed. Using: {', '.join(trimmed_list)}."
                )
            invalid_list = st.session_state.get("_graph_invalid", [])
            if invalid_list and not st.session_state.get("_graph_fixed_commas"):
                shown = ", ".join(invalid_list[:10]) + (
                    "…" if len(invalid_list) > 10 else ""
                )
                st.warning(
                    f"Some entries are not valid HEX colors and were ignored: {shown}. "
                    "Please use #RRGGBB, e.g., #12295D."
                )

            if st.session_state.graph_colors:
                st.write("Selected Colors:")
                selected_html = "<div style='display:flex;flex-wrap:wrap;'>"
                for color in st.session_state.graph_colors:
                    selected_html += (
                        f"<div title='{color}' style='background:{color};width:40px;height:40px;"
                        "margin:2px;border:1px solid #ccc;border-radius:4px;'></div>"
                    )
                selected_html += "</div>"
                st.markdown(selected_html, unsafe_allow_html=True)

    # ---------------- Excel Header Color ----------------
    with right:
        st.markdown(
            "<div style='font-size:18px; font-weight:600; margin-bottom:4px;'>Choose Color for Excel Formatting</div>",
            unsafe_allow_html=True,
        )
        tab_choice = st.radio(
            "",
            ["Use Default Color", "Pick Your Own"],
            key="tab_choice",
            label_visibility="collapsed",
        )

        import re

        def _is_hex(s: str) -> bool:
            return bool(re.fullmatch(r"\#[0-9A-Fa-f]{6}", (s or "").strip()))

        def _all_hex(s: str):
            return [m.upper() for m in re.findall(r"\#[0-9A-Fa-f]{6}", s or "")]

        def _invalid_tokens(s: str):
            parts = [p.strip() for p in re.split(r"[,\\\s]+", s or "") if p.strip()]
            bad = []
            for p in parts:
                if _is_hex(p):
                    continue
                if _all_hex(p):
                    continue
                bad.append(p)
            return bad

        if "tab_color" not in st.session_state:
            st.session_state.tab_color = ""
        if "tab_input" not in st.session_state:
            st.session_state.tab_input = ""
        if "_prev_tab_choice" not in st.session_state:
            st.session_state._prev_tab_choice = tab_choice
        if "_tab_suppress_once" not in st.session_state:
            st.session_state._tab_suppress_once = False
        if "_tab_invalid" not in st.session_state:
            st.session_state._tab_invalid = False
        if "_tab_trimmed" not in st.session_state:
            st.session_state._tab_trimmed = False
        if "_tab_invalid_tokens" not in st.session_state:
            st.session_state._tab_invalid_tokens = []
        if "_tab_empty" not in st.session_state:
            st.session_state._tab_empty = False
        if "_clear_tab_now" not in st.session_state:
            st.session_state._clear_tab_now = False

        if st.session_state._clear_tab_now:
            st.session_state.tab_color = ""
            st.session_state.tab_input = ""
            st.session_state._tab_invalid = False
            st.session_state._tab_trimmed = False
            st.session_state._tab_invalid_tokens = []
            st.session_state._tab_empty = True
            st.session_state._clear_tab_now = False

        if st.session_state._prev_tab_choice != tab_choice:
            st.session_state._prev_tab_choice = tab_choice
            st.session_state._tab_invalid = False
            st.session_state._tab_trimmed = False
            st.session_state._tab_invalid_tokens = []
            st.session_state._tab_empty = True
            st.session_state._tab_suppress_once = tab_choice == "Pick Your Own"

        if tab_choice == "Use Default Color":
            st.session_state.tab_color = default_tab_color
            st.write("Using default excel color:")
            st.markdown(
                f"<div style='background:{default_tab_color};width:40px;height:40px;"
                "margin:2px;border:1px solid #ccc;border-radius:4px;'></div>",
                unsafe_allow_html=True,
            )
        else:

            def update_tab_color():
                raw = (st.session_state.tab_input or "").strip()
                matches = _all_hex(raw)
                if matches:
                    first = matches[0].upper()
                    st.session_state.tab_color = first
                    st.session_state.tab_input = first
                    st.session_state._tab_invalid = False
                    st.session_state._tab_trimmed = (raw.upper() != first) or (
                        len(matches) > 1
                    )
                    st.session_state._tab_invalid_tokens = []
                    st.session_state._tab_empty = False
                else:
                    if raw == "":
                        st.session_state.tab_color = ""
                        st.session_state.tab_input = ""
                        st.session_state._tab_invalid = False
                        st.session_state._tab_trimmed = False
                        st.session_state._tab_invalid_tokens = []
                        st.session_state._tab_empty = True
                    else:
                        st.session_state.tab_color = ""
                        st.session_state._tab_invalid = True
                        st.session_state._tab_trimmed = False
                        st.session_state._tab_invalid_tokens = _invalid_tokens(raw)
                        st.session_state._tab_empty = False

            # -------- ROW 1: swatch (left) + Set Tab Color (right) on the SAME row --------
            st.markdown(
                "<div style='font-size:14px; font-weight:600; margin-top:4px;'>Pick Excel Header Color</div>",
                unsafe_allow_html=True,
            )
            tc_sw, tc_btn = st.columns([6, 4])

            with tc_sw:
                tab_color_picker = st.color_picker(
                    label="",
                    value=st.session_state.tab_color or "#12295D",
                    key="tab_color_picker",
                    label_visibility="collapsed",
                )

            with tc_btn:
                if st.button(
                    "Set Header Color",
                    key="set_tab_color_btn",
                    type="primary",
                    use_container_width=True,
                ):
                    picked = (tab_color_picker or "").strip().upper()
                    if _is_hex(picked):
                        st.session_state.tab_color = picked
                        st.session_state.tab_input = picked
                        st.session_state._tab_invalid = False
                        st.session_state._tab_trimmed = False
                        st.session_state._tab_invalid_tokens = []
                        st.session_state._tab_empty = False
                    else:
                        st.session_state.tab_color = ""
                        st.session_state.tab_input = ""
                        st.session_state._tab_invalid = False
                        st.session_state._tab_trimmed = False
                        st.session_state._tab_invalid_tokens = []
                        st.session_state._tab_empty = True

            # -------- ROW 2: label + input (left) and Clear (right), perfectly aligned --------
            ti_input, ti_clear = st.columns([7, 2])

            with ti_input:
                st.text_input(
                    "Insert HEX Color (choose only single color)",
                    key="tab_input",
                    on_change=update_tab_color,
                    placeholder="#AABBCC",
                    help="Enter a 7-character HEX color like #AABBCC.",
                )

            def _set_clear_tab_flag():
                st.session_state._clear_tab_now = True

            with ti_clear:
                st.markdown("<div style='height:28px'></div>", unsafe_allow_html=True)
                st.button(
                    "Clear",
                    key="clear_tab_input_btn",
                    type="primary",
                    use_container_width=True,
                    on_click=_set_clear_tab_flag,
                )

            # -------- messages & preview --------
            raw_tab = (st.session_state.tab_input or "").strip()
            suppress_tab = st.session_state._tab_suppress_once
            if raw_tab and not suppress_tab:
                if st.session_state.get("_tab_trimmed"):
                    st.warning(
                        f"Extra characters or multiple colors were removed. Using {st.session_state.tab_color}."
                    )
                elif st.session_state.get("_tab_invalid") and not st.session_state.get(
                    "_tab_empty"
                ):
                    toks = st.session_state.get("_tab_invalid_tokens", [])
                    if toks:
                        shown = ", ".join(toks[:10]) + ("…" if len(toks) > 10 else "")
                        st.warning(
                            f"Entered HEX color is not valid and were ignored: {shown}. "
                            "Please use #RRGGBB, e.g., #12295D."
                        )
                    else:
                        st.warning(
                            "Entered HEX color is not valid and were ignored. Please use #RRGGBB, e.g., #12295D."
                        )

            if raw_tab == "" or suppress_tab:
                st.session_state._tab_invalid = False
                st.session_state._tab_trimmed = False
                st.session_state._tab_invalid_tokens = []
                st.session_state._tab_empty = True
                st.session_state._tab_suppress_once = False

            excel_pick = st.session_state.get("tab_choice") == "Pick Your Own"
            show_tab_preview = (
                not excel_pick and bool(st.session_state.tab_color)
            ) or (excel_pick and raw_tab != "" and _is_hex(st.session_state.tab_color))
            if show_tab_preview:
                st.write("Selected Color:")
                st.markdown(
                    f"<div style='background:{st.session_state.tab_color};width:40px;height:40px;"
                    "margin:2px;border:1px solid #ccc;border-radius:4px;'></div>",
                    unsafe_allow_html=True,
                )

    # ------------------- Color Customization Section ----------- #

    # ------------------------- Proceed ------------------------- #

    st.markdown("<div style='height:14px'></div>", unsafe_allow_html=True)

    import re

    def _is_hex(s):
        return bool(re.fullmatch(r"\#[0-9A-Fa-f]{6}", (s or "").strip()))

    graph_pick = st.session_state.get("graph_choice") == "Pick Your Own"
    excel_pick = st.session_state.get("tab_choice") == "Pick Your Own"

    graph_provided = (
        (len(st.session_state.get("graph_colors", [])) > 0) if graph_pick else True
    )
    excel_input_nonempty = bool((st.session_state.get("tab_input", "") or "").strip())
    excel_provided = (
        (_is_hex(st.session_state.get("tab_color", "")) and excel_input_nonempty)
        if excel_pick
        else True
    )

    need_proceed = graph_pick or excel_pick
    can_proceed = graph_provided and excel_provided

    if "proceed_confirmed" not in st.session_state:
        st.session_state.proceed_confirmed = False

    prev_sig = st.session_state.get("_color_state_sig")
    state_sig = (
        graph_pick,
        excel_pick,
        tuple(st.session_state.get("graph_colors", [])),
        st.session_state.get("tab_color", ""),
    )

    if prev_sig != state_sig:
        if need_proceed:
            st.session_state.proceed_confirmed = False
        st.session_state["_color_state_sig"] = state_sig

    if need_proceed:
        if not can_proceed:
            if not graph_provided and not excel_provided:
                msg = "Please choose colors for Graph and Excel Header to proceed."
            elif not graph_provided:
                msg = "Please choose at least one HEX color for Graph to proceed."
            else:
                msg = "Please choose a HEX color for Excel Header to proceed."
            st.warning(msg)

        proceed_clicked = st.button(
            "Proceed",
            disabled=not can_proceed,
            key="proceed_colors",
            use_container_width=True,
            type="primary",
        )
        if proceed_clicked:
            st.session_state.proceed_confirmed = True

    # ------------------------- Proceed ------------------------- #
    
    # ------------------------- Export -------------------------- #   
    # st.markdown("#### Export")

    # ppt_choice = st.radio(
    #     "Do you want to generate PPT Deck?", 
    #     options=["No", "Yes"], 
    #     index=0, 
    #     horizontal=True, 
    #     key="eda_generate_ppt_choice",
    # )

    # ppt_template_bytes = None
    # if ppt_choice == "Yes":
    #     ppt_up = st.file_uploader(
    #         "**Upload PPTX template**",
    #         type=["pptx"],
    #         key="eda_ppt_template_uploader",
    #     )
    #     if ppt_up:
    #         ppt_template_bytes = ppt_up.getvalue()


    # params["graph_colors"] = st.session_state.get("graph_colors", [])
    # params["tab_color"] = st.session_state.get("tab_color", "")

    # missing = []
    # if (metric_var_sel == "(None)") and not same_col:
    #     missing.append("Metric")
    # if (value_var_sel == "(None)") and not same_col:
    #     missing.append("Value")

    # export_enabled_colors = (not need_proceed) or (st.session_state.get("proceed_confirmed") and can_proceed)
    # export_enabled = export_enabled_colors and (len(missing) == 0)

    # if len(missing) > 0:
    #     pretty = " and ".join(missing)
    #     st.warning(f"Please choose the {pretty} column(s) before exporting.")

    # if export_enabled:
    #     with st.spinner("Preparing output..."):
    #         excel_bytes = build_eda_excel_bytes(
    #             file_bytes=up.getvalue(), file_name=up.name, params=params
    #         )
    #         if st.session_state.get("eda_generate_ppt_choice") == "Yes":
    #             ppt_bytes = build_eda_ppt_bytes(
    #                 file_bytes=up.getvalue(), file_name=up.name, params=params, template_bytes=ppt_template_bytes
    #             )
    #             # Bundle ZIP
    #             import io, zipfile
    #             zip_buffer = io.BytesIO()
    #             with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
    #                 zf.writestr("EDA_Final_Output.xlsx", excel_bytes)
    #                 zf.writestr("EDA_Deck.pptx", ppt_bytes)
    #             zip_buffer.seek(0)
    #             st.download_button(
    #                 label="🚀 Generate & Download (Excel + PPT ZIP)",
    #                 data=zip_buffer.getvalue(),
    #                 file_name="EDA_Output.zip",
    #                 mime="application/zip",
    #                 key="eda_generate_zip_download"
    #             )
    #         else:
    #             st.download_button(
    #                 label="🚀 Generate & Download EDA Workbook",
    #                 data=excel_bytes,
    #                 file_name="EDA_Final_Output.xlsx",
    #                 mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    #                 key="eda_generate_and_download"
    #             )
    # else:
    #     if not export_enabled_colors:
    #         st.caption("Set required colors and click Proceed to enable Export button.")

    
    st.markdown("#### Export")

    left, right = st.columns([1,1])

    with left:
        st.markdown('<div class="align-top">', unsafe_allow_html=True)
        st.markdown("###### Choose files to download")
        want_excel = st.checkbox(
            "Generate Excel Workbook",
            value=True,
            key="export_excel"
        )

        want_ppt = st.checkbox(
            "Generate PPT Deck",
            value=True,
            key="export_ppt"
        )

        st.markdown("</div>", unsafe_allow_html=True)

    with right:
        st.markdown('<div class="align-top">', unsafe_allow_html=True)
        st.markdown("###### Upload PPTX template")
        ppt_up = st.file_uploader(
            "",
            type=["pptx"],
            key="eda_ppt_template_uploader"
        )

        st.markdown("</div>", unsafe_allow_html=True)

    ppt_template_bytes = ppt_up.getvalue() if ppt_up else None

    params["graph_colors"] = st.session_state.get("graph_colors", [])
    params["tab_color"] = st.session_state.get("tab_color", "")

    missing = []
    if (metric_var_sel == "(None)") and not same_col:
        missing.append("Metric")
    if (value_var_sel == "(None)") and not same_col:
        missing.append("Value")

    need_ppt_template = want_ppt and (ppt_template_bytes is None)
    at_least_one = want_excel or want_ppt

    export_enabled_colors = (not need_proceed) or (
        st.session_state.get("proceed_confirmed") and can_proceed
    )

    export_enabled = (
        at_least_one
        and not need_ppt_template
        and (len(missing) == 0)
        and export_enabled_colors
    )

    if len(missing) > 0:
        st.warning(f"Please choose the {', '.join(missing)} column(s).")

    if want_ppt and ppt_template_bytes is None:
        st.warning("Please upload a PPTX template to enable export.")

    if not export_enabled_colors:
        st.caption("Set required colors and click Proceed to enable Export button.")

    if not at_least_one:
        st.info("Select at least one output type.")

    excel_bytes = b""
    ppt_bytes = b""

    if export_enabled:
        with st.spinner("Preparing output..."):

            if want_excel:
                excel_bytes = build_eda_excel_bytes(
                    file_bytes=up.getvalue(),
                    file_name=up.name,
                    params=params
                )

            if want_ppt:
                ppt_bytes = build_eda_ppt_bytes(
                    file_bytes=up.getvalue(),
                    file_name=up.name,
                    params=params,
                    template_bytes=ppt_template_bytes
                )

    if want_excel and want_ppt:
        import io, zipfile
        zip_buffer = io.BytesIO()

        if export_enabled:
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                zf.writestr("EDA_Final_Output.xlsx", excel_bytes)
                zf.writestr("EDA_Deck.pptx", ppt_bytes)
            zip_buffer.seek(0)

        st.download_button(
            "🚀 Generate & Download (Excel + PPT ZIP)",
            data=zip_buffer.getvalue() if export_enabled else b"",
            file_name="EDA_Output.zip",
            mime="application/zip",
            disabled=not export_enabled
        )

    elif want_excel:
        st.download_button(
            "🚀 Generate & Download Excel Workbook",
            data=excel_bytes if export_enabled else b"",
            file_name="EDA_Final_Output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            disabled=not export_enabled
        )

    elif want_ppt:
        st.download_button(
            "🚀 Generate & Download PPT Deck",
            data=ppt_bytes if export_enabled else b"",
            file_name="EDA_Deck.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            disabled=not export_enabled
        )
    # ------------------------- Export ------------------------- #
