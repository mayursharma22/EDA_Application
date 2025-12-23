import streamlit as st
import pandas as pd
import io
import zipfile
import re
from datetime import datetime
from pandas.api.types import is_categorical_dtype
from pandas.api.types import is_datetime64_any_dtype
from pandas.api.types import is_numeric_dtype

# ------------------------------------------------------------
# -------------------------- Session State -------------------
# ------------------------------------------------------------
# Rename entries are scoped: base(df) vs melted
if "rename_entries_df" not in st.session_state:
    st.session_state.rename_entries_df = {}  # {file_key: [ {"column","new_name"} ]}
if "rename_entries_melted" not in st.session_state:
    st.session_state.rename_entries_melted = {}  # {file_key: [ {"column","new_name"} ]}
if "schema_entries" not in st.session_state:
    st.session_state.schema_entries = {}  # {file_key: [ {"name","default"} ]}
if "file_index" not in st.session_state:
    st.session_state.file_index = 0
if "grouped_data" not in st.session_state:
    st.session_state.grouped_data = {}  # single-file grouped outputs
if "page" not in st.session_state:
    st.session_state.page = "process"
if "preprocessed_data" not in st.session_state:
    st.session_state.preprocessed_data = {}  # per-file processed base data
if "combined_df" not in st.session_state:
    st.session_state.combined_df = pd.DataFrame()
if "final_melted" not in st.session_state:
    st.session_state.final_melted = pd.DataFrame()
if "final_grouped" not in st.session_state:
    st.session_state.final_grouped = pd.DataFrame()
if "multi_mode" not in st.session_state:
    st.session_state.multi_mode = False
if "column_mappings_multi" not in st.session_state:
    st.session_state.column_mappings_multi = []
<<<<<<<< HEAD:src/eda_app/eda_data_processing.py
# Raw data + per-file settings
if "raw_data" not in st.session_state:
    st.session_state.raw_data = {}  # {file_key: raw_df}
if "date_settings" not in st.session_state:
    st.session_state.date_settings = {}  # {file_key: {"date_col": str, "new_name": str}}
if "channel_settings" not in st.session_state:
    st.session_state.channel_settings = {}  # {file_key: {mode/cols/values}}
# Breakdown state
if "breakdown_state" not in st.session_state:
========
if 'raw_data' not in st.session_state:
    st.session_state.raw_data = {}            # {file_key: raw_df}
if 'date_settings' not in st.session_state:
    st.session_state.date_settings = {}       # {file_key: {"date_col": str, "new_name": str}}
if 'channel_settings' not in st.session_state:
    st.session_state.channel_settings = {}    # {file_key: {mode/cols/values}}
if 'breakdown_state' not in st.session_state:
>>>>>>>> 9c072067ba049488615f6356de28fb2fb62e6523:src/eda_data_processing.py
    st.session_state.breakdown_state = {}

st.title("📊 Data Preparation for EDA")

# ------------------------------------------------------------
# --------------------- Constants / Helpers ------------------
# ------------------------------------------------------------
META_COL = "_source_file"  # hidden metadata column carrying the source filename


def to_display_df(df: pd.DataFrame) -> pd.DataFrame:
    """Hide metadata column from previews."""
    return df.drop(columns=[META_COL], errors="ignore")


def ensure_default(key: str, value):
    if key not in st.session_state:
        st.session_state[key] = value


def safe_filename(text: str) -> str:
    text = str(text).strip()
    if text == "" or text.lower() == "nan":
        text = "Blank"
    return re.sub(r'[\\/*?:"<>|\n]+', "_", text)


def load_file_to_df(uploaded):
    name = uploaded.name
    if name.endswith(".csv"):
        df = pd.read_csv(uploaded)
    else:
        df = pd.read_excel(uploaded, engine="openpyxl")
    df[META_COL] = name
    df["_source_file_display"] = safe_filename(name)
    return df

<<<<<<<< HEAD:src/eda_app/eda_data_processing.py
========
# def is_effectively_numeric(series, threshold=0.99):
#     if pd.api.types.is_datetime64_any_dtype(series):
#         return False
#     coerced = pd.to_numeric(series, errors='coerce')
#     if len(series) == 0:
#         return False
#     non_numeric_ratio = coerced.isna().sum() / len(series)
#     return non_numeric_ratio < threshold
>>>>>>>> 9c072067ba049488615f6356de28fb2fb62e6523:src/eda_data_processing.py

def is_effectively_numeric(series, threshold=0.99):
    if pd.api.types.is_datetime64_any_dtype(series):
        return False
<<<<<<<< HEAD:src/eda_app/eda_data_processing.py
    coerced = pd.to_numeric(series, errors="coerce")
    if len(series) == 0:
        return False
    non_numeric_ratio = coerced.isna().sum() / len(series)
========

    if series.dtype == 'O':
        s_clean = _preclean_numeric_strings(series)
        mask = s_clean != ""
        if not mask.any():
            return False
        coerced = pd.to_numeric(s_clean[mask], errors='coerce')
        non_numeric_ratio = coerced.isna().sum() / int(mask.sum())
    else:
        mask = series.notna()
        if not mask.any():
            return False
        coerced = pd.to_numeric(series[mask], errors='coerce')
        non_numeric_ratio = coerced.isna().sum() / int(mask.sum())

>>>>>>>> 9c072067ba049488615f6356de28fb2fb62e6523:src/eda_data_processing.py
    return non_numeric_ratio < threshold


EXCLUDE_MELT_PREFIXES = ["dma_code", "dma_name", "unnamed", "npi", "hcp"]


def filter_melt_candidates(candidates):
    lc_prefixes = tuple(p.lower() for p in EXCLUDE_MELT_PREFIXES)

    def keep(col):
        c = str(col).strip().lower()
        return not (c in lc_prefixes or c.startswith(lc_prefixes))

    return [c for c in candidates if keep(c)]


def decat_for_display(df: pd.DataFrame) -> pd.DataFrame:
    """
    Return a copy where categorical columns are converted back to object.
    This prevents errors when we do .fillna("") or write CSVs.
    """
    out = df.copy()
    for c in out.columns:
        if is_categorical_dtype(out[c]):
            out[c] = out[c].astype(object)
    return out


# ------------------------------------------------------------
# ------------------- Date / Channel UIs ---------------------
# ------------------------------------------------------------


def date_selection_ui(df, file_key):
    def ensure_default(key, value):
        if key not in st.session_state:
            st.session_state[key] = value

    def parse_dates_series(s):
        """Robust date parsing for mixed formats."""
        if is_datetime64_any_dtype(s):
            parsed = pd.to_datetime(s, errors="coerce")
            return parsed.dt.normalize(), 1.0 - parsed.isna().mean()

        s = s.astype(str).str.strip()
        p1 = pd.to_datetime(
            s, errors="coerce", infer_datetime_format=True, dayfirst=True
        )
        p2 = pd.to_datetime(
            s, errors="coerce", infer_datetime_format=True, dayfirst=False
        )
        parsed = p1 if p1.isna().mean() <= p2.isna().mean() else p2

        if parsed.isna().mean() > 0.20:
            formats = [
                "%Y-%m-%d",
                "%y-%m-%d",
                "%d-%m-%Y",
                "%m-%d-%Y",
                "%d/%m/%Y",
                "%m/%d/%Y",
                "%Y/%m/%d",
                "%d.%m.%Y",
                "%m.%d.%Y",
            ]
            parsed = parsed.copy()
            for fmt in formats:
                mask = parsed.isna()
                if not mask.any():
                    break
                try:
                    parsed.loc[mask] = pd.to_datetime(
                        s[mask], format=fmt, errors="coerce"
                    )
                except Exception:
                    pass

        parsed = pd.to_datetime(parsed, errors="coerce").dt.normalize()
        success_ratio = 1.0 - parsed.isna().mean()
        return parsed, success_ratio

    st.write("📅 **Date Column Selection**")
    st.session_state.setdefault("date_settings", {})
    saved = st.session_state.date_settings.get(file_key, {})

    date_columns = [col for col in df.columns if is_datetime64_any_dtype(df[col])]

    if not date_columns:
        candidate_scores = {}
        for col in df.columns:
            if df[col].dtype == "O":
                sample = df[col].dropna().astype(str).head(50)
                parsed = pd.to_datetime(
                    sample, errors="coerce", infer_datetime_format=True
                )
                success_ratio = 1.0 - parsed.isna().mean()
                if success_ratio >= 0.8:
                    candidate_scores[col] = success_ratio
        if candidate_scores:
            date_columns = sorted(
                candidate_scores, key=candidate_scores.get, reverse=True
            )

    all_columns = list(df.columns)

    saved_choice = saved.get("date_col")
    if saved_choice in all_columns:
        default_date = saved_choice
    elif date_columns:
        default_date = date_columns[0]
    else:
        default_date = "None"

    ensure_default(f"date_col_{file_key}", default_date)
    ensure_default(f"new_date_col_{file_key}", saved.get("new_name", ""))

    cols = st.columns([1, 2])
    with cols[0]:
        options = ["None"] + all_columns
        date_col = st.selectbox(
            "Select date column", options, key=f"date_col_{file_key}"
        )
    with cols[1]:
        new_date_column = st.text_input(
            "Rename date column (Optional)",
            key=f"new_date_col_{file_key}",
            placeholder="e.g. WeekStartDate",
        )

    if date_col != "None":
        col_dtype = df[date_col].dtype
        if not is_datetime64_any_dtype(df[date_col]):
            parsed, success = parse_dates_series(df[date_col])

            if success >= 0.80:
                st.success(
                    f"Conversion successful for '{date_col}' (YYYY-MM-DD). Parsed: {success:.0%} of values."
                )
            else:
                st.warning(
                    f"⚠️ Please select correct date column. Conversion for '{date_col}' parsed only {success:.0%}."
                )
        else:
            st.info(
                f"'{date_col}' is already a datetime column. Normalized for preview."
            )

    st.session_state.date_settings[file_key] = {
        "date_col": date_col,
        "new_name": new_date_column,
    }

    return date_col, new_date_column

<<<<<<<< HEAD:src/eda_app/eda_data_processing.py

# def channel_selection_ui(df, file_key):
#     st.write("📦 **Channel Column Selection**")
#     saved = st.session_state.channel_settings.get(file_key, {})
#     ensure_default(f"channel_radio_{file_key}", saved.get("mode", "Yes"))
#     channel_available = st.radio("Is there a 'Channel' column?", ("Yes", "No"), key=f"channel_radio_{file_key}")
#     final_channel_name = None
#     if channel_available == "Yes":
#         ensure_default(f"channel_col_{file_key}", saved.get("channel_col", df.columns[0] if len(df.columns) else ""))
#         ensure_default(f"new_channel_col_{file_key}", saved.get("new_name", ""))
#         cols = st.columns([1, 2])
#         with cols[0]:
#             channel_col = st.selectbox("Select channel column", df.columns, key=f"channel_col_{file_key}")
#         with cols[1]:
#             new_channel_name = st.text_input("Rename channel column (Optional)",
#                                              key=f"new_channel_col_{file_key}",
#                                              placeholder="e.g. SourceChannel")
#         final_channel_name = new_channel_name if new_channel_name else channel_col
#         st.session_state.channel_settings[file_key] = {"mode": "Yes", "channel_col": channel_col, "new_name": new_channel_name}
#     else:
#         ensure_default(f"custom_channel_col_{file_key}", saved.get("custom_name", ""))
#         ensure_default(f"channel_value_{file_key}", saved.get("value", ""))
#         cols = st.columns([1, 2])
#         with cols[0]:
#             custom_channel_col = st.text_input("Enter custom channel column name",
#                                                key=f"custom_channel_col_{file_key}",
#                                                placeholder="e.g. SourceChannel")
#         with cols[1]:
#             channel_value = st.text_input("Enter value for channel column",
#                                           key=f"channel_value_{file_key}",
#                                           placeholder="e.g. Google Ads")
#         final_channel_name = custom_channel_col if custom_channel_col else None
#         st.session_state.channel_settings[file_key] = {"mode": "No", "custom_name": custom_channel_col, "value": channel_value}
#     return final_channel_name


========
>>>>>>>> 9c072067ba049488615f6356de28fb2fb62e6523:src/eda_data_processing.py
def channel_selection_ui(df, file_key):
    import streamlit as st
    from pandas.api.types import is_numeric_dtype, is_datetime64_any_dtype

    def ensure_default(key, value):
        if key not in st.session_state:
            st.session_state[key] = value

    st.write("📦 **Channel Column Selection**")
    saved = st.session_state.channel_settings.get(file_key, {})
    date_cfg = st.session_state.date_settings.get(file_key, {})
    orig_date_col = date_cfg.get("date_col", None)
    new_date_col = (date_cfg.get("new_name") or "").strip() or None

    if new_date_col and new_date_col in df.columns:
        active_date_col = new_date_col
    elif orig_date_col and orig_date_col in df.columns:
        active_date_col = orig_date_col
    else:
        active_date_col = None

    META_EXCLUDE = {"_source_file", "_source_file_display"}
    eligible_columns = [
        col
        for col in df.columns
        if col not in META_EXCLUDE
        and (active_date_col is None or col != active_date_col)
        and not is_numeric_dtype(df[col])
        and not is_datetime64_any_dtype(df[col])
    ]

    channel_like = next(
        (col for col in eligible_columns if str(col).strip().lower() == "channel"), None
    )

    saved_choice = saved.get("channel_col")
    if saved_choice in eligible_columns:
        default_channel = saved_choice
    elif channel_like:
        default_channel = channel_like
    else:
        default_channel = "None"

    ensure_default(f"channel_radio_{file_key}", saved.get("mode", "Yes"))
    channel_available = st.radio(
        "Is there a 'Channel' column?", ("Yes", "No"), key=f"channel_radio_{file_key}"
    )
    final_channel_name = None

    if channel_available == "Yes":
        ensure_default(f"channel_col_{file_key}", default_channel)
        ensure_default(f"new_channel_col_{file_key}", saved.get("new_name", ""))

        cols = st.columns([1, 2])
        with cols[0]:
            options = ["None"] + eligible_columns
            channel_col = st.selectbox(
                "Select channel column", options, key=f"channel_col_{file_key}"
            )
        with cols[1]:
            new_channel_name = st.text_input(
                "Rename channel column (Optional)",
                key=f"new_channel_col_{file_key}",
                placeholder="e.g. SourceChannel",
            )

        final_channel_name = new_channel_name if new_channel_name else channel_col
        st.session_state.channel_settings[file_key] = {
            "mode": "Yes",
            "channel_col": channel_col,
            "new_name": new_channel_name,
        }

    else:
        ensure_default(f"custom_channel_col_{file_key}", saved.get("custom_name", ""))
        ensure_default(f"channel_value_{file_key}", saved.get("value", ""))

        cols = st.columns([1, 2])
        with cols[0]:
            custom_channel_col = st.text_input(
                "Enter custom channel column name",
                key=f"custom_channel_col_{file_key}",
                placeholder="e.g. SourceChannel",
            )
        with cols[1]:
            channel_value = st.text_input(
                "Enter value for channel column",
                key=f"channel_value_{file_key}",
                placeholder="e.g. Google Ads",
            )

        final_channel_name = custom_channel_col if custom_channel_col else None
        st.session_state.channel_settings[file_key] = {
            "mode": "No",
            "custom_name": custom_channel_col,
            "value": channel_value,
        }

    return final_channel_name


# ------------------------------------------------------------
# --------------------- Scoped Rename Logic ------------------
# ------------------------------------------------------------
def _get_rename_store(scope: str):
    return (
        st.session_state.rename_entries_melted
        if scope == "melted"
        else st.session_state.rename_entries_df
    )


def apply_renames(df: pd.DataFrame, file_key: str, scope: str = "df"):
    """Apply rename map to df (in-place) using the correct scope (base or melted)."""
    store = _get_rename_store(scope)
    entries = store.get(file_key, [])
    rename_map = {
        e.get("column", ""): e.get("new_name", "")
        for e in entries
        if e.get("column") and e.get("new_name")
    }
    if rename_map:
        df.rename(columns=rename_map, inplace=True)


def _sync_rename_entry_from_widgets(file_key: str, i: int, scope: str):
    store = _get_rename_store(scope)
    entry = store[file_key][i]
    col_key = f"rename_col_{scope}_{file_key}_{i}"
    name_key = f"rename_name_{scope}_{file_key}_{i}"
    entry["column"] = st.session_state.get(col_key, entry.get("column", ""))
    entry["new_name"] = st.session_state.get(name_key, entry.get("new_name", ""))


def rename_other_columns_ui(df: pd.DataFrame, file_key: str, scope: str = "df"):
    """
    Stable, scoped UI for renaming columns.
    scope: 'df' (base)  |  'melted' (post-melt)
    """
    st.write("✏️ **Rename Other Columns**")
    store = _get_rename_store(scope)
    if file_key not in store:
        store[file_key] = []
    for i, entry in enumerate(store[file_key]):
        cols = st.columns([1, 2, 0.1])
        with cols[0]:
            current_columns = [c for c in df.columns if c != META_COL]
            default_col = entry.get(
                "column", current_columns[0] if current_columns else ""
            )
            ensure_default(f"rename_col_{scope}_{file_key}_{i}", default_col)
            st.selectbox(
                "Column",
                current_columns or [""],
                key=f"rename_col_{scope}_{file_key}_{i}",
                on_change=_sync_rename_entry_from_widgets,
                args=(file_key, i, scope),
            )
        with cols[1]:
            ensure_default(
                f"rename_name_{scope}_{file_key}_{i}", entry.get("new_name", "")
            )
            st.text_input(
                "New name",
                key=f"rename_name_{scope}_{file_key}_{i}",
                placeholder="e.g. FinalColumnName",
                on_change=_sync_rename_entry_from_widgets,
                args=(file_key, i, scope),
            )
        with cols[2]:
            if st.button(
                "🗑️",
                key=f"del_rename_{scope}_{file_key}_{i}",
                help="Delete rename entry",
            ):
                store[file_key].pop(i)
                st.rerun()
        _sync_rename_entry_from_widgets(file_key, i, scope)

    if st.button("➕ Add Rename Entry", key=f"add_rename_{scope}_{file_key}"):
        store[file_key].append({"column": "", "new_name": ""})
        st.rerun()


# ------------------------------------------------------------
# -------------------------- Schema --------------------------
# ------------------------------------------------------------
def add_schema_ui(df, file_key):
    st.write("🧋 **Add New Columns**")
    if file_key not in st.session_state.schema_entries:
        st.session_state.schema_entries[file_key] = []
    for i in range(len(st.session_state.schema_entries[file_key])):
        cols = st.columns([2, 2, 0.1])
        with cols[0]:
            ensure_default(
                f"schema_entry_{file_key}_{i}",
                st.session_state.schema_entries[file_key][i].get("name", ""),
            )
            schema_name = st.text_input(
                "Schema column",
                key=f"schema_entry_{file_key}_{i}",
                placeholder="e.g. FinalColumnName",
            )
        with cols[1]:
            ensure_default(
                f"default_value_{file_key}_{i}",
                st.session_state.schema_entries[file_key][i].get("default", ""),
            )
            default_value = st.text_input(
                "Default value (optional)",
                key=f"default_value_{file_key}_{i}",
                placeholder="e.g. N/A",
            )
        with cols[2]:
            if st.button(
                "🗑️", key=f"del_schema_{file_key}_{i}", help="Delete schema entry"
            ):
                st.session_state.schema_entries[file_key].pop(i)
                st.rerun()
        if i < len(st.session_state.schema_entries[file_key]):
            st.session_state.schema_entries[file_key][i] = {
                "name": schema_name,
                "default": default_value,
            }

    if st.button("➕ Add Schema Column", key=f"add_schema_{file_key}"):
        st.session_state.schema_entries[file_key].append({"name": "", "default": ""})
        st.rerun()


# ------------------------------------------------------------
# ---------------------- Multi-file helpers ------------------
# ------------------------------------------------------------
def apply_column_mappings(df, mappings):
    """Merge/replace multiple columns into a unified column; never drop META_COL."""
    if not mappings:
        return df
    new_df = df.copy()
    for mapping in mappings:
        selected_cols = mapping.get("columns", [])
        unified_name = mapping.get("unified_name", "")
        if not selected_cols or not unified_name:
            continue
        for col in selected_cols:
            if col in new_df.columns:
                if unified_name in new_df.columns:
                    new_df[unified_name] = new_df[unified_name].combine_first(
                        new_df[col]
                    )
                else:
                    new_df[unified_name] = new_df[col]
        for col in selected_cols:
            if col != unified_name and col in new_df.columns and col != META_COL:
                new_df.drop(columns=[col], inplace=True)
    return new_df

<<<<<<<< HEAD:src/eda_app/eda_data_processing.py

def detect_numeric_candidates_across_files(dfs, threshold=0.1):
    """(Legacy) Full scan; we will prefer cached light preview below."""
========
def _preclean_numeric_strings(s: pd.Series) -> pd.Series:
    """
    Make numeric-like strings parseable by pd.to_numeric:
    - Trim whitespace
    - Convert parentheses to negatives: '(123)' -> '-123'
    - Strip common currency symbols
    - Remove thousands separators (comma, non-breaking/thin spaces)
    - Remove any remaining internal whitespace
    """
    currency_symbols_pattern = r'[\$\₹£€¥₩₽฿₪₫₴₦₺]'
    s = s.astype(str).str.strip()
    s = s.str.replace(r'^\((.*)\)$', r'-\1', regex=True)
    s = s.str.replace(currency_symbols_pattern, '', regex=True)  
    s = s.str.replace(r'[,\u00A0\u2009\u202F]', '', regex=True)    
    s = s.str.replace(r'\s+', '', regex=True)                     
    return s

def detect_numeric_candidates_across_files(
    dfs, threshold: float = 0.99, meta_col: str | None = None):
    """
    Detect columns that are 'numeric enough' (invalid ratio <= threshold) in ANY DataFrame.
    Also converts those columns to numeric in-place in each DataFrame.

    Returns:
        sorted list of unique candidate column names.
    """
>>>>>>>> 9c072067ba049488615f6356de28fb2fb62e6523:src/eda_data_processing.py
    candidates = set()

    for df in dfs:
        for col in df.columns:
            if meta_col is not None and col == meta_col:
                continue
            s = df[col]
            if pd.api.types.is_datetime64_any_dtype(s):
                continue
<<<<<<<< HEAD:src/eda_app/eda_data_processing.py
            mask = (s.astype(str).str.strip() != "") if s.dtype == "O" else s.notna()
            if mask.any():
                coerced = pd.to_numeric(s[mask], errors="coerce")
                invalid = coerced.isna().sum()
                total = int(mask.sum())
                if total > 0 and (invalid / total) < threshold:
                    candidates.add(col)
========
            if s.dtype == 'O':
                s_clean = _preclean_numeric_strings(s)
                mask = s_clean != ""
                if not mask.any():
                    continue
                coerced = pd.to_numeric(s_clean[mask], errors='coerce')
            else:
                mask = s.notna()
                if not mask.any():
                    continue
                coerced = pd.to_numeric(s[mask], errors='coerce')
            invalid = int(coerced.isna().sum())
            total = int(mask.sum())
            invalid_ratio = (invalid / total) if total > 0 else 1.0
            if total > 0 and invalid_ratio <= threshold:
                candidates.add(col)
                if s.dtype == 'O':
                    df[col] = pd.to_numeric(_preclean_numeric_strings(s), errors='coerce')
                else:
                    df[col] = pd.to_numeric(s, errors='coerce')
>>>>>>>> 9c072067ba049488615f6356de28fb2fb62e6523:src/eda_data_processing.py
    return sorted(candidates)


def map_numeric_candidates_to_unified(candidates, mappings):
    if not mappings:
        return sorted(set(candidates))
    out = set(candidates)
    for m in mappings:
        cols = set(m.get("columns", []))
        uni = m.get("unified_name", "")
        if not cols or not uni:
            continue
        if out.intersection(cols):
            out = (out - cols) | {uni}
    return sorted(out)


def build_light_preview_for_numeric(pre_list, sample_rows: int = 500):
    """
    For each file, only look at up to N rows per column and compute invalid numeric ratio.
    Returns: list of dicts {col: {"invalid": int, "total": int}}
    """
    previews = []
    for df in pre_list:
        small = df if (sample_rows is None or sample_rows <= 0) else df.head(sample_rows)
        meta = {}
        for col in small.columns:
            if col == META_COL or pd.api.types.is_datetime64_any_dtype(small[col]):
                continue
            s = small[col]
<<<<<<<< HEAD:src/eda_app/eda_data_processing.py
            mask = (s.astype(str).str.strip() != "") if s.dtype == "O" else s.notna()
            if mask.any():
                coerced = pd.to_numeric(s[mask], errors="coerce")
                invalid = int(coerced.isna().sum())
                total = int(mask.sum())
                meta[col] = {"invalid": invalid, "total": total}
========
            if s.dtype == 'O':
                s_clean = _preclean_numeric_strings(s)
                mask = s_clean != ""
                if not mask.any():
                    continue
                coerced = pd.to_numeric(s_clean[mask], errors='coerce')
            else:
                mask = s.notna()
                if not mask.any():
                    continue
                coerced = pd.to_numeric(s[mask], errors='coerce')
            invalid = int(coerced.isna().sum())
            total   = int(mask.sum())
            meta[col] = {"invalid": invalid, "total": total}

>>>>>>>> 9c072067ba049488615f6356de28fb2fb62e6523:src/eda_data_processing.py
        previews.append(meta)
    return previews


def collapse_preview_counts(previews):
    counts = {}
    for meta in previews:
        for col, m in meta.items():
            if col not in counts:
                counts[col] = {"tot": 0, "invalid": 0}
            counts[col]["tot"] += m["total"]
            counts[col]["invalid"] += m["invalid"]
    return counts


def counts_to_signature(counts: dict) -> tuple:
    return tuple(sorted((col, v["tot"], v["invalid"]) for col, v in counts.items()))


@st.cache_data(show_spinner=False)
<<<<<<<< HEAD:src/eda_app/eda_data_processing.py
def cached_numeric_candidates(
    counts_sig: tuple, exclude_prefixes: tuple, threshold: float = 0.1
) -> list:
========
def cached_numeric_candidates(counts_sig: tuple,
                              exclude_prefixes: tuple,
                              threshold: float = 0.99) -> list:
>>>>>>>> 9c072067ba049488615f6356de28fb2fb62e6523:src/eda_data_processing.py
    out = []
    lc_prefixes = tuple(p.lower() for p in exclude_prefixes)

    def keep(col):
        c = str(col).strip().lower()
        return not (c in lc_prefixes or c.startswith(lc_prefixes))

    for col, tot, invalid in counts_sig:
        if not keep(col):
            continue
        if tot > 0:
            numeric_count = tot - invalid
            if ((invalid / tot) < threshold) or (numeric_count >= 1):
                out.append(col)

    return sorted(out)


<<<<<<<< HEAD:src/eda_app/eda_data_processing.py
========

>>>>>>>> 9c072067ba049488615f6356de28fb2fb62e6523:src/eda_data_processing.py
# ------------------------------------------------------------
# ----------------------- Per-file transforms ----------------
# ------------------------------------------------------------
def apply_date_transform(df: pd.DataFrame, file_key: str) -> pd.DataFrame:
    cfg = st.session_state.date_settings.get(file_key, {})
    date_col = cfg.get("date_col")
    new_name = cfg.get("new_name", "")
    if not date_col or date_col not in df.columns:
        return df
    out = df.copy()
    out[date_col] = pd.to_datetime(out[date_col], errors="coerce")
    if new_name:
        out[new_name] = out[date_col]
        out.drop(columns=[date_col], inplace=True)
    return out


def apply_channel_transform(df: pd.DataFrame, file_key: str) -> pd.DataFrame:
    cfg = st.session_state.channel_settings.get(file_key, {})
    if not cfg:
        return df
    out = df.copy()
    if cfg.get("mode") == "Yes":
        col = cfg.get("channel_col", "")
        new_name = cfg.get("new_name", "")
        if col and col in out.columns:
            if new_name:
                out[new_name] = out[col]
                out.drop(columns=[col], inplace=True)
    else:
        custom = cfg.get("custom_name", "")
        value = cfg.get("value", "")
        if custom:
            out[custom] = value
    return out


def build_processed_df(base_df: pd.DataFrame, file_key: str) -> pd.DataFrame:
    df = base_df.copy()
    if META_COL not in df.columns:
        df[META_COL] = file_key
    df = apply_date_transform(df, file_key)
    df = apply_channel_transform(df, file_key)
    apply_renames(df, file_key, scope="df")
    for schema in st.session_state.schema_entries.get(file_key, []):
        col = schema.get("name", "")
        default_val = schema.get("default", "")
        if col and col not in df.columns:
            df[col] = default_val
    if META_COL not in df.columns:
        df[META_COL] = file_key
    return df


def save_current_file(file_key: str):
    base_df = st.session_state.raw_data[file_key]
    processed = build_processed_df(base_df, file_key)
    st.session_state.preprocessed_data[file_key] = processed


# ------------------------------------------------------------
# ------------------- Breakdown File helpers -----------------
# ------------------------------------------------------------
def blanks_per_file(df: pd.DataFrame, col: str) -> pd.DataFrame:
    if col not in df.columns:
        return pd.DataFrame(columns=[META_COL, "blank_count"])
    s = df[col].astype(str)
    mask_blank = df[col].isna() | (s.str.strip() == "")
    if META_COL in df.columns:
        grp = (
            df.loc[mask_blank]
            .groupby(META_COL, dropna=False)
            .size()
            .reset_index(name="blank_count")
        )
        return grp.sort_values("blank_count", ascending=False)
    count = int(mask_blank.sum())
    return pd.DataFrame({META_COL: ["(single file)"], "blank_count": [count]})


def apply_per_file_replacements(
    df: pd.DataFrame, col: str, per_file_map: dict
) -> pd.DataFrame:
    if col not in df.columns or not per_file_map:
        return df
    out = df.copy()
    is_blank = out[col].isna() | (out[col].astype(str).str.strip() == "")
    for fname, replacement in per_file_map.items():
        if replacement is None:
            continue
        rep = str(replacement).strip()
        if rep == "":
            continue
        mask = (out[META_COL].astype(str) == str(fname)) & is_blank
        if mask.any():
            if is_categorical_dtype(out[col]) and rep not in list(
                out[col].cat.categories
            ):
                out[col] = out[col].cat.add_categories([rep])
            out.loc[mask, col] = rep
    return out


def recompute_grouped_from_melt_base(
    melt_base: pd.DataFrame, group_cols: list
) -> pd.DataFrame:
    grp_base = melt_base.dropna(subset=["Values"])
    if grp_base.empty or not group_cols:
        return pd.DataFrame()
    dims = [c for c in group_cols if c in grp_base.columns]
    for c in dims:
        if c in grp_base.columns and grp_base[c].dtype == "O":
            grp_base[c] = grp_base[c].astype("category")
    grouped_df = (
        grp_base.groupby(dims, dropna=False, observed=True)["Values"]
        .sum()
        .reset_index()
    )
    for c in dims:
        if is_categorical_dtype(grouped_df[c]):
            grouped_df[c] = grouped_df[c].astype(object)
    return grouped_df


def as_categorical(df: pd.DataFrame, dims: list) -> pd.DataFrame:
    out = df
    for c in dims:
        if c in out.columns and out[c].dtype == "O":
            out[c] = out[c].astype("category")
    return out


def build_pre_group_base(melt_base: pd.DataFrame, group_columns: list) -> pd.DataFrame:
    """
    Pre-aggregate melted data at group_columns + [_source_file] to keep per-file traceability
    but drastically shrink size before blank replacement & final group.
    """
    if melt_base is None or not len(melt_base):
        return pd.DataFrame()
    cols = [c for c in group_columns if c in melt_base.columns]
    if META_COL not in melt_base.columns or not cols:
        return pd.DataFrame()
    gcols = cols + [META_COL]
    base = melt_base.dropna(subset=["Values"])
    base = as_categorical(base, [c for c in gcols if c != META_COL])
    pre = base.groupby(gcols, dropna=False, observed=True)["Values"].sum().reset_index()
    for c in gcols:
        if c in pre.columns and is_categorical_dtype(pre[c]):
            pre[c] = pre[c].astype(object)
    return pre


def apply_blank_replacements_and_final_group(
    pre_group_base: pd.DataFrame, bcol: str, group_columns: list, per_file_map: dict
) -> pd.DataFrame:
    """
    Apply per-file blank replacements on the already pre-aggregated base,
    then collapse _source_file and return final grouped result.
    """
    if pre_group_base.empty or bcol not in pre_group_base.columns:
        return pd.DataFrame()
    out = pre_group_base.copy()
    blank_mask = out[bcol].isna() | (out[bcol].astype(str).str.strip() == "")
    if per_file_map:
        map_series = out.loc[blank_mask, META_COL].map(
            lambda f: str(per_file_map.get(f, "")).strip()
        )
        fill_mask = blank_mask & map_series.ne("")
        if fill_mask.any():
            if is_categorical_dtype(out[bcol]):
                new_vals = pd.Index(map_series[fill_mask].dropna().unique())
                to_add = [
                    v for v in new_vals if v not in list(out[bcol].cat.categories)
                ]
                if to_add:
                    out[bcol] = out[bcol].cat.add_categories(to_add)
            out.loc[fill_mask, bcol] = map_series[fill_mask]
    dims = [c for c in group_columns if c in out.columns]
    for c in dims:
        if out[c].dtype == "O":
            out[c] = out[c].astype("category")
    final = out.groupby(dims, dropna=False, observed=True)["Values"].sum().reset_index()
    for c in dims:
        if is_categorical_dtype(final[c]):
            final[c] = final[c].astype(object)
    return final


def render_download_and_breakdown_melt(
    final_df: pd.DataFrame,
    melt_base: pd.DataFrame,
    group_columns: list,
    default_csv_name: str,
    state_key: str,
):
    """
    Melt-only panel:
    - Preview final_df
    - Breakdown radio (No/Yes)
    - If blanks exist for chosen breakdown column, collect per-file replacements (gate download)
    - Recompute using pre-aggregated base if needed; provide single ZIP button in breakdown mode
    """
    st.markdown("#### 🧾 Final Output")
    display_df = decat_for_display(final_df).fillna("")
    st.dataframe(to_display_df(display_df).head(10))

    ################ New added code for go on prev page from final page ######
    if st.session_state.get("page") == "combine" and st.session_state.get("multi_mode", False):
        total_files = len(st.session_state.get("raw_data", {}))
        if total_files > 0:
            if st.button("⬅️ Save & Previous", key=f"final_prev_only__{state_key}", use_container_width=True):
                st.session_state.file_index = total_files - 1
                st.session_state.page = "process"
                st.rerun()

    ################ New added code for go on prev page from final page ######

    ensure_default(f"breakdown_toggle_{state_key}", "No")
    breakdown_choice = st.radio(
        "Breakdown final output by a specific column?",
        ("No", "Yes"),
        horizontal=True,
        key=f"breakdown_toggle_{state_key}",
    )

    if breakdown_choice == "No":
        st.download_button(
            "⬇️ Download Final Data",
            data=display_df.to_csv(index=False).encode("utf-8"),
            file_name=default_csv_name,
            mime="text/csv",
            key=f"download_single_{state_key}",
        )
        return

    breakdown_cols = [
        c
        for c in final_df.columns
        if c not in (META_COL, "_source_file_display", "Values")
    ]
    if not breakdown_cols:
        st.warning("No eligible columns available to breakdown.")
        return

    ensure_default(f"breakdown_col_{state_key}", breakdown_cols[0])
    bcol = st.selectbox(
        "Select the column to split the final output:",
        options=breakdown_cols,
        key=f"breakdown_col_{state_key}",
    )

    pre_key = f"pre_group_base__{state_key}__{'|'.join(group_columns)}"
    if pre_key not in st.session_state:
        st.session_state[pre_key] = build_pre_group_base(melt_base, group_columns)
    pre_group_base = st.session_state[pre_key]

    def get_blank_summary_and_gate(selected_col: str):
        if pre_group_base is None or selected_col not in pre_group_base.columns:
            return pd.DataFrame(columns=[META_COL, "blank_count"]), False
        s = pre_group_base[selected_col].astype(str)
        mask_blank = pre_group_base[selected_col].isna() | (s.str.strip() == "")
        grp = (
            pre_group_base.loc[mask_blank]
            .groupby(META_COL, dropna=False)
            .size()
            .reset_index(name="blank_count")
        )
        grp = grp.sort_values("blank_count", ascending=False)
        needs_save = (not grp.empty) and (grp["blank_count"].sum() > 0)
        return grp, needs_save

    blank_summary, needs_save = get_blank_summary_and_gate(bcol)
    saved_flag_key = f"saved_flag_{state_key}_{bcol}"
    ensure_default(saved_flag_key, False)

    if needs_save and not st.session_state.get(saved_flag_key, False):
        st.info("The selected column has blanks for the following source files:")
        st.dataframe(blank_summary)

        rep_key = f"replacements_{state_key}_{bcol}"
        if rep_key not in st.session_state.breakdown_state:
            st.session_state.breakdown_state[rep_key] = {}

        st.write("🧩 Provide a name for blanks (per file):")
        for _, row in blank_summary.iterrows():
            fname = row[META_COL]
            ensure_default(f"repl_{state_key}_{bcol}_{fname}", "")
            st.session_state.breakdown_state[rep_key][fname] = st.text_input(
                f"Blank label for file: {fname}",
                key=f"repl_{state_key}_{bcol}_{fname}",
                placeholder="Enter replacement for blanks in this file",
            )

        inputs_map = st.session_state.breakdown_state.get(rep_key, {})
        all_filled = len(inputs_map) > 0 and all(
            str(v).strip() for v in inputs_map.values()
        )
        if st.button(
            "💾 Save Replacements",
            key=f"save_repl_{state_key}_{bcol}",
            disabled=not all_filled,
        ):
            st.session_state[saved_flag_key] = True
            st.success("Saved replacements.")
        if not st.session_state.get(saved_flag_key, False):
            return

    per_file_map = st.session_state.breakdown_state.get(
        f"replacements_{state_key}_{bcol}", {}
    )
    if per_file_map or needs_save:
        work_df = apply_blank_replacements_and_final_group(
            pre_group_base, bcol, group_columns, per_file_map
        )
    else:
        work_df = final_df

    if work_df is None or work_df.empty:
        st.warning("No data after recomputation.")
        return

    work_df_disp = decat_for_display(work_df).fillna("")

    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        values = (
            work_df_disp[bcol]
            .astype(str)
            .replace({"": "Blank"})
            .fillna("Blank")
            .unique()
            .tolist()
        )
        for val in values:
            part = work_df_disp[
                work_df_disp[bcol].astype(str).replace({"": "Blank"}).fillna("Blank")
                == val
            ]
            csv_bytes = to_display_df(part).to_csv(index=False).encode("utf-8")
            zf.writestr(f"{safe_filename(val)}_Processed_File.csv", csv_bytes)
    zip_buffer.seek(0)

    st.download_button(
        "⬇️ Download Breakdown (ZIP)",
        data=zip_buffer,
        file_name=f"Breakdown_{safe_filename(bcol)}_{datetime.today().strftime('%Y%m%d')}.zip",
        mime="application/zip",
        key=f"download_zip_{state_key}",
    )


# ------------------------------------------------------------
# ------------------------- PAGE 1: PROCESS ------------------
# ------------------------------------------------------------
if st.session_state.page == "process":
<<<<<<<< HEAD:src/eda_app/eda_data_processing.py
    uploaded_files = st.file_uploader(
        "**Upload CSV or Excel files**",
        type=["csv", "xlsx"],
        accept_multiple_files=True,
    )
========
    uploaded_files = st.file_uploader("**Upload CSV or Excel files**",
                                      type=["csv", "xlsx"],
                                      accept_multiple_files=True)
    
    ################ New added code for go on prev page from final page ######
    if (not uploaded_files) and st.session_state.raw_data:
        class _PseudoFile:
            def __init__(self, name): self.name = name
        uploaded_files = [_PseudoFile(n) for n in st.session_state.raw_data.keys()]
    ################ New added code for go on prev page from final page ######


>>>>>>>> 9c072067ba049488615f6356de28fb2fb62e6523:src/eda_data_processing.py
    if uploaded_files:
        total_files = len(uploaded_files)
        st.session_state.multi_mode = total_files > 1

        current_index = st.session_state.file_index
        current_file = uploaded_files[current_index]
        file_key = current_file.name

        if file_key not in st.session_state.raw_data:
            st.session_state.raw_data[file_key] = load_file_to_df(current_file)
        base_df = st.session_state.raw_data[file_key]

        st.subheader(
            f"Processing File {current_index + 1} of {total_files}: {file_key}"
        )
        st.write("📄 **Preview of uploaded data**:")
        st.dataframe(to_display_df(base_df).head())

        # Init per-file stores
        if file_key not in st.session_state.schema_entries:
            st.session_state.schema_entries[file_key] = []
        if file_key not in st.session_state.rename_entries_df:
            st.session_state.rename_entries_df[file_key] = []
        if file_key not in st.session_state.rename_entries_melted:
            st.session_state.rename_entries_melted[file_key] = []

        # Render date/channel UIs
        date_selection_ui(base_df, file_key)
        channel_selection_ui(base_df, file_key)

        # ---------- SINGLE-FILE FLOW  ----------
        if not st.session_state.multi_mode:
            rename_other_columns_ui(base_df, file_key, scope="df")
            add_schema_ui(base_df, file_key)

            st.write("🧪 **Melt Columns (Long Format)**")

            date_cfg = st.session_state.date_settings.get(file_key, {})
            orig_dc = date_cfg.get("date_col", None)
            new_dc = (date_cfg.get("new_name") or "").strip() or None

            if new_dc and new_dc in base_df.columns:
                active_date_col = new_dc
            elif orig_dc and orig_dc in base_df.columns:
                active_date_col = orig_dc
            else:
                active_date_col = None

            # numeric_candidates = []
            # df_for_detect = base_df.copy()
            # for col in df_for_detect.columns:
            #     if col == META_COL:
            #         continue
            #     if active_date_col and col == active_date_col:
            #         continue
            #     if is_effectively_numeric(df_for_detect[col]):
            #         df_for_detect[col] = pd.to_numeric(df_for_detect[col], errors='coerce')
            #         numeric_candidates.append(col)
            # numeric_candidates = filter_melt_candidates(numeric_candidates)

            
            numeric_candidates = []
            df_for_detect = base_df.copy()

            for col in df_for_detect.columns:
                # Skip metadata and the active date column
                if col == META_COL:
                    continue
                if active_date_col and col == active_date_col:
                    continue
<<<<<<<< HEAD:src/eda_app/eda_data_processing.py
                if is_effectively_numeric(df_for_detect[col]):
                    df_for_detect[col] = pd.to_numeric(
                        df_for_detect[col], errors="coerce"
                    )
========

                s = df_for_detect[col]

                # Count how many values are parseable as numeric (full column, no sampling)
                if s.dtype == 'O':
                    s_clean = _preclean_numeric_strings(s)
                    mask = s_clean != ""
                    if mask.any():
                        coerced = pd.to_numeric(s_clean[mask], errors='coerce')
                        numeric_count = int(coerced.notna().sum())
                    else:
                        numeric_count = 0
                else:
                    mask = s.notna()
                    if mask.any():
                        coerced = pd.to_numeric(s[mask], errors='coerce')
                        numeric_count = int(coerced.notna().sum())
                    else:
                        numeric_count = 0

                # Accept if either:
                # 1) the column is "effectively numeric" (your function), OR
                # 2) it has at least one numeric value after cleaning (fallback)
                if is_effectively_numeric(s) or (numeric_count >= 1):
>>>>>>>> 9c072067ba049488615f6356de28fb2fb62e6523:src/eda_data_processing.py
                    numeric_candidates.append(col)

            # Keep your existing prefix exclusions
            numeric_candidates = filter_melt_candidates(numeric_candidates)


            metric_cols_key = f"metric_cols_single__{file_key}"
            ensure_default(metric_cols_key, [])
            st.session_state.selected_metric_columns_single = st.session_state[
                metric_cols_key
            ]

            metric_columns = st.multiselect(
                "Select Metric columns to melt", numeric_candidates, key=metric_cols_key
            )
            st.session_state.selected_metric_columns_single = metric_columns

            clean_map_key = f"clean_name_map__{file_key}"
            if clean_map_key not in st.session_state:
                st.session_state[clean_map_key] = {}

            for c in metric_columns:
                st.session_state[clean_map_key].setdefault(c, c)

            if metric_columns:
                st.write("Provide clean names for selected metrics (optional)")
                for col in metric_columns:
                    current_mapped = st.session_state[clean_map_key].get(col, col)
                    new_name = st.text_input(
                        f"Clean name for '{col}'",
                        key=f"clean_single_{file_key}__{col}",
                        value="",
                        placeholder=f"e.g. {current_mapped}",
                    )
                    if new_name.strip():
                        st.session_state[clean_map_key][col] = new_name.strip()

            if st.button("Apply Melt", key=f"apply_melt_single__{file_key}"):
                df = build_processed_df(base_df, file_key)
                melt_base = df.copy() 
                for mc in metric_columns:
                    if mc in melt_base.columns:
<<<<<<<< HEAD:src/eda_app/eda_data_processing.py
                        melt_base[mc] = pd.to_numeric(melt_base[mc], errors="coerce")
========
                        s = melt_base[mc]
                        if s.dtype == 'O':
                            melt_base[mc] = pd.to_numeric(_preclean_numeric_strings(s), errors='coerce')
                        else:
                            melt_base[mc] = pd.to_numeric(s, errors='coerce')

>>>>>>>> 9c072067ba049488615f6356de28fb2fb62e6523:src/eda_data_processing.py

                melted_df = melt_base.melt(
                    id_vars=[c for c in melt_base.columns if c not in metric_columns],
                    value_vars=metric_columns,
                    var_name="MetricsOrig",
                    value_name="Values",
                )

                cmap = st.session_state[clean_map_key]
                melted_df["Metrics"] = (
                    melted_df["MetricsOrig"]
                    .map(lambda x: cmap.get(x, x))
                    .astype("category")
                )

                st.session_state[f"single_melted_df__{file_key}"] = melted_df
                st.session_state[f"single_melt_applied__{file_key}"] = True
                st.success("Melt applied successfully.")

            melted_key = f"single_melted_df__{file_key}"
            if st.session_state.get(melted_key, pd.DataFrame()).empty is False:
                _df = st.session_state[melted_key]
                cmap = st.session_state.get(clean_map_key, {})
                if cmap:
                    new_labels = _df["MetricsOrig"].map(lambda x: cmap.get(x, x))
                    if not new_labels.equals(_df["Metrics"].astype(str)):
                        _df = _df.copy()
                        _df["Metrics"] = pd.Categorical(new_labels)
                        st.session_state[melted_key] = _df

            melt_applied_key = f"single_melt_applied__{file_key}"
            if (
                st.session_state.get(melt_applied_key)
                and not st.session_state.get(melted_key, pd.DataFrame()).empty
            ):
                melted_df = st.session_state[melted_key]
                st.write("🧮 **Final Output**")

                group_cols_key = f"group_cols_{file_key}"
                ensure_default(group_cols_key, [])

                with st.form(f"group_form_{file_key}"):
                    group_columns = st.multiselect(
                        "Select columns to include in final output",
                        [
                            c
                            for c in melted_df.columns
                            if c
                            not in (
                                "Values",
                                META_COL,
                                "_source_file_display",
                                "MetricsOrig",
                            )
                        ],
                        key=group_cols_key,
                    )
                    do_group = st.form_submit_button("🚀 Group Data")

                if do_group:
                    if not group_columns:
                        st.error("Please select at least one column to group by.")
                    else:
                        grp_base = melted_df.dropna(subset=["Values"])
                        if grp_base.empty:
                            st.warning(
                                "All `Values` are NaN after melt. Check metric selection and numeric parsing."
                            )
                            st.session_state.grouped_data[file_key] = pd.DataFrame()
                        else:
                            dims = [c for c in group_columns if c in grp_base.columns]
                            for c in dims:
                                if grp_base[c].dtype == "O":
                                    grp_base[c] = grp_base[c].astype("category")
                            grouped_df = (
                                grp_base.groupby(dims, dropna=False, observed=True)[
                                    "Values"
                                ]
                                .sum()
                                .reset_index()
                            )
                            from pandas.api.types import is_categorical_dtype

                            for c in dims:
                                if is_categorical_dtype(grouped_df[c]):
                                    grouped_df[c] = grouped_df[c].astype(object)
                            st.session_state.grouped_data[file_key] = grouped_df

                if (
                    file_key in st.session_state.grouped_data
                    and not st.session_state.grouped_data[file_key].empty
                ):
                    export_df = st.session_state.grouped_data[file_key]
                    selected_groups = st.session_state[group_cols_key]

                    render_download_and_breakdown_melt(
                        final_df=export_df,
                        melt_base=melted_df,
                        group_columns=selected_groups,
                        default_csv_name="Processed_Data.csv",
                        state_key=f"single_melt_{file_key}",
                    )

        # ---------- MULTI-FILE FLOW ----------
        else:
            rename_other_columns_ui(base_df, file_key, scope="df")
            add_schema_ui(base_df, file_key)

            df_preview = build_processed_df(base_df, file_key)
            st.write("🧾 **Preview after per‑file transformations (pre‑melt):**")
            st.dataframe(to_display_df(df_preview).head())

            nav_prev, nav_next = st.columns([1, 1])
            with nav_prev:
                if st.button(
                    "⬅️ Save & Previous",
                    key=f"prev_file_{file_key}",
                    use_container_width=True,
                    disabled=(current_index == 0),
                ):
                    save_current_file(file_key)
                    st.session_state.file_index -= 1
                    st.rerun()
            with nav_next:
                if st.button(
                    "➡️ Save & Next",
                    key=f"next_file_{file_key}",
                    use_container_width=True,
                ):
                    save_current_file(file_key)
                    if (st.session_state.file_index + 1) < total_files:
                        st.session_state.file_index += 1
                        st.rerun()
                    else:
                        st.session_state.page = "combine"
                        st.rerun()

# ------------------------------------------------------------
# --------------------- PAGE 2: COMBINE (MELT ONLY) ----------
# ------------------------------------------------------------
elif st.session_state.page == "combine":
    if not st.session_state.multi_mode:
        all_grouped = list(st.session_state.grouped_data.values())
        if not all_grouped:
            st.warning(
                "No grouped data available. Please go back and process the file."
            )
            st.stop()
        combined_df = pd.concat(all_grouped, axis=0, ignore_index=True).fillna("")
        st.session_state.combined_df = combined_df
        selected_columns = st.multiselect(
            "**Select columns to include in final download:**",
            [
                c
                for c in combined_df.columns
                if c not in (META_COL, "_source_file_display")
            ],
        )
        filtered_df = combined_df[selected_columns] if selected_columns else combined_df
        st.write("🔎 **Filtered Preview:**")
        st.dataframe(to_display_df(filtered_df).head(5))
        csv = (
            decat_for_display(filtered_df)
            .fillna("")
            .to_csv(index=False)
            .encode("utf-8")
        )
        st.download_button(
            "⬇️ Download Final Data",
            data=csv,
            file_name="Processed_Combined_Data.csv",
            mime="text/csv",
        )
    else:
        if not st.session_state.preprocessed_data:
            st.warning("No preprocessed files found. Please go back and process files.")
            st.stop()

        pre_list = list(st.session_state.preprocessed_data.values())
        combined_df = pd.concat(pre_list, axis=0, ignore_index=True, sort=False)

        # ---------- Column Harmonization ----------
        st.write("🧩 **Column Harmonization** (optional)")

        def update_mapping_multi(index):
            st.session_state.column_mappings_multi[index]["columns"] = (
                st.session_state.get(f"merge_cols_multi_{index}", [])
            )
            st.session_state.column_mappings_multi[index]["unified_name"] = (
                st.session_state.get(f"unified_name_multi_{index}", "")
            )

        for i, mapping in enumerate(st.session_state.column_mappings_multi):
            cols = st.columns([2, 2, 0.1])
            with cols[0]:
                st.multiselect(
                    f"Select columns to merge (Mapping {i + 1})",
                    combined_df.columns.tolist(),
                    default=mapping.get("columns", []),
                    key=f"merge_cols_multi_{i}",
                    on_change=update_mapping_multi,
                    args=(i,),
                )
            with cols[1]:
                st.text_input(
                    "Unified column name",
                    value=mapping.get("unified_name", ""),
                    key=f"unified_name_multi_{i}",
                    placeholder="e.g. Campaign",
                    on_change=update_mapping_multi,
                    args=(i,),
                )
            with cols[2]:
                if st.button(
                    "🗑️", key=f"del_mapping_multi_{i}", help="Delete column mapping"
                ):
                    st.session_state.column_mappings_multi.pop(i)
                    st.rerun()

        if st.button("➕ Add Column Mapping", key="add_column_mapping_multi"):
            st.session_state.column_mappings_multi.append(
                {"columns": [], "unified_name": ""}
            )
            st.rerun()

        harmonized_df = apply_column_mappings(
            combined_df, st.session_state.column_mappings_multi
        )

        # ---------- Melt on Combined/Harmonized Data ----------
        st.write("🧪 **Melt Columns (across all files)**")
        light_preview = build_light_preview_for_numeric(pre_list, sample_rows=None)
        counts = collapse_preview_counts(light_preview)
        counts_sig = counts_to_signature(counts)
<<<<<<<< HEAD:src/eda_app/eda_data_processing.py
        cands = cached_numeric_candidates(
            counts_sig, tuple(EXCLUDE_MELT_PREFIXES), threshold=0.1
        )
        cands = map_numeric_candidates_to_unified(
            cands, st.session_state.column_mappings_multi
        )
========
        cands = cached_numeric_candidates(counts_sig, tuple(EXCLUDE_MELT_PREFIXES), threshold=0.99)
        cands = map_numeric_candidates_to_unified(cands, st.session_state.column_mappings_multi)
>>>>>>>> 9c072067ba049488615f6356de28fb2fb62e6523:src/eda_data_processing.py
        cands = [c for c in cands if c in harmonized_df.columns]
        st.session_state.cached_numeric_candidates_multi = cands
        numeric_candidates = st.session_state.get("cached_numeric_candidates_multi", [])
        numeric_candidates = filter_melt_candidates(numeric_candidates)
        st.success(f"Found {len(numeric_candidates)} candidate metric columns.")

        metric_columns = st.multiselect(
            "Select Metric columns to melt",
            numeric_candidates,
            key="metric_cols_combined",
        )
        st.session_state.selected_metric_columns = metric_columns
        if metric_columns:
            st.write("Provide clean names for selected metrics (optional)")
            clean_names = []
            for col in metric_columns:
                clean_names.append(
                    st.text_input(
                        f"Clean name for '{col}'",
                        key=f"clean_combined_{col}",
                        placeholder=f"e.g. {col}",
                    )
                    or col
                )

            if st.button("Apply Melt", key="apply_melt_multi"):
                melt_base = harmonized_df.copy()  
                for mc in metric_columns:
<<<<<<<< HEAD:src/eda_app/eda_data_processing.py
                    melt_base[mc] = pd.to_numeric(melt_base[mc], errors="coerce")
========
                    s = melt_base[mc]
                    if s.dtype == 'O':
                        melt_base[mc] = pd.to_numeric(_preclean_numeric_strings(s), errors='coerce')
                    else:
                        melt_base[mc] = pd.to_numeric(s, errors='coerce')

>>>>>>>> 9c072067ba049488615f6356de28fb2fb62e6523:src/eda_data_processing.py
                melted_df = melt_base.melt(
                    id_vars=[c for c in melt_base.columns if c not in metric_columns],
                    value_vars=metric_columns,
                    var_name="Metrics",
                    value_name="Values",
                )
                metric_name_map = dict(zip(metric_columns, clean_names))
                melted_df["Metrics"] = (
                    melted_df["Metrics"].map(metric_name_map).astype("category")
                )
                st.session_state.final_melted = melted_df
                st.success("Melt applied successfully.")

        if not st.session_state.final_melted.empty:
            melted_df = st.session_state.final_melted
            st.write("🧮 **Final Output**")
            group_columns = st.multiselect(
                "Select columns to include in final output",
                [
                    c
                    for c in melted_df.columns
                    if c not in ("Values", META_COL, "_source_file_display")
                ],
                key="group_cols_combined",
            )

            if st.button("🚀 Group Data", key="group_data_combined"):
                if not group_columns:
                    st.error("Please select at least one column to group by.")
                else:
                    grp_base = melted_df.dropna(subset=["Values"])
                    if grp_base.empty:
                        st.warning(
                            "All `Values` are NaN after melt. Check metric selection/harmonization/number parsing."
                        )
                        st.session_state.final_grouped = pd.DataFrame()
                    else:
                        dims = [c for c in group_columns if c in grp_base.columns]
                        for c in dims:
                            if grp_base[c].dtype == "O":
                                grp_base[c] = grp_base[c].astype("category")
                        final_grouped = (
                            grp_base.groupby(dims, dropna=False, observed=True)[
                                "Values"
                            ]
                            .sum()
                            .reset_index()
                        )
                        for c in dims:
                            if is_categorical_dtype(final_grouped[c]):
                                final_grouped[c] = final_grouped[c].astype(object)
                        st.session_state.final_grouped = final_grouped

            if not st.session_state.final_grouped.empty:
                export_df = st.session_state.final_grouped
                render_download_and_breakdown_melt(
                    final_df=export_df,
                    melt_base=melted_df,
                    group_columns=group_columns,
                    default_csv_name="Processed_Combined_Data.csv",
                    state_key="multi_melt_combined",
                )

        else:
            st.info("Select metric columns and click **Apply Melt** to proceed.")
