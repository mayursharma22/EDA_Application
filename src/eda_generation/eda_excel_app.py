# Standard Library
import os
import tempfile
import hashlib

# Third Party imports
import pandas as pd
import streamlit as st

# Internal imports
from .eda_excel_generation import run as eda_excel_run


st.set_page_config(page_title="EDA Generation", layout="wide")


# ---------- helpers ----------
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


# ---------- main ----------
def eda_generation():
    st.title("📈 EDA Generation")
    st.caption(
        "Upload a CSV and configure parameters. Generates a formatted Excel workbook."
    )

    # Upload
    up = st.file_uploader("**Upload CSV**", type=["csv"], key="eda_uploader")
    if not up:
        st.info("Upload a CSV file to begin.")
        return

    df = pd.read_csv(up)
    st.markdown("#### Preview:")
    st.dataframe(df.head())
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

    default_metric_idx = (
        list(df.columns).index("Metrics") if "Metrics" in df.columns else 0
    )
    default_value_idx = (
        list(df.columns).index("Values") if "Values" in df.columns else 0
    )

    metric_var = st.selectbox(
        "Metric column",
        options=df.columns,
        index=default_metric_idx,
        key="eda_metric_var",
    )
    value_var = st.selectbox(
        "Value column", options=df.columns, index=default_value_idx, key="eda_value_var"
    )

    metric_names = (
        df[metric_var].dropna().astype(str).str.strip().unique().tolist()
        if metric_var in df.columns
        else []
    )
    cost_metric_options = ["(None)"] + sorted(metric_names)
    cost_metric = st.selectbox(
        "Cost/Spend metric (optional)",
        options=cost_metric_options,
        index=0,
        key="eda_cost_metric",
    )
    cost_var = "" if cost_metric == "(None)" else cost_metric

    st.markdown("#### Select Breakdown Fields")

    dim_exclude = {date_var, metric_var, value_var}
    dim_candidates = [c for c in df.columns if c not in dim_exclude]

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

    st.markdown("#### Export")
    params["graph_colors"] = st.session_state.get("graph_colors", [])
    params["tab_color"] = st.session_state.get("tab_color", "")

    export_enabled = (not need_proceed) or (
        st.session_state.get("proceed_confirmed") and can_proceed
    )

    if export_enabled:
        with st.spinner("Preparing workbook..."):
            excel_bytes = build_eda_excel_bytes(
                file_bytes=up.getvalue(), file_name=up.name, params=params
            )
    else:
        excel_bytes = b""

    st.download_button(
        label="🚀 Generate & Download EDA Workbook",
        data=excel_bytes if export_enabled else b"",
        file_name="EDA_Final_Output.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="eda_generate_and_download",
        disabled=not export_enabled,
    )

    if not export_enabled:
        st.caption("Set required colors and click Proceed to enable Export button.")

    # ------------------------- Export ------------------------- #
