# --------------------------------------------------------------
# eda_generation.py
# Author: Saurabh Shinkar
# Version: 1.0
# Description: Automated EDA generation for performance data with weekly/quarterly pivots and summary.
# Date: 2025-10-28
# --------------------------------------------------------------
import numpy as np
import pandas as pd
import string
import time
import os
import re
import itertools
from pathlib import Path
from pandas.tseries.offsets import QuarterEnd

# ------------------------------------------------------------------
# Excel helper constants
# ------------------------------------------------------------------
Color_Palette = [
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

# Color_Palette = [
#     '#6DBB4F', '#9997C9', '#84BAE5', '#16FFA7', '#FFE67A',
#     '#AEAEBC', '#A2F9FB', '#00CACF', '#E5E5E9', '#FFF3CD',
#     '#41547D', '#FFAA00', '#2B49A6', '#FFDC69', '#06757E',
#     '#996DDF', '#FF644C', '#F23A1D', '#5B19C4', '#12295D'
# ]

Tab_Color_Palette = "#12295D"

df_columns = list(string.ascii_uppercase) + [
    f + s for f in string.ascii_uppercase for s in string.ascii_uppercase
]

# ------------------------------------------------------------------
# Helper: shorten sheet name
# ------------------------------------------------------------------
_short_counter = 0
_short_map = {}


def short_name(full: str) -> str:
    global _short_counter
    if full not in _short_map:
        base = full[:20] + (full[-8:] if len(full) > 28 else "")
        candidate = base[:31]
        orig = candidate
        while candidate in _short_map.values():
            _short_counter += 1
            suffix = str(_short_counter)[-3:]
            candidate = orig[: 31 - len(suffix)] + suffix
        _short_map[full] = candidate
    return _short_map[full]


def pretty_col(name: str) -> str:
    words = name.replace("_", " ").split()
    updated = [w.upper() if len(w) <= 3 else w.title() for w in words]
    return " ".join(updated)


# ------------------------------------------------------------------
# MAIN FUNCTION
# ------------------------------------------------------------------
def run(params: dict):
    csv_path = Path(params["csv_path"])
    df = pd.read_csv(csv_path)

    date_var = params.get("date_var", "date").strip()
    date_grain = (
        params.get("date_grain", "weekly").strip().lower()
    )  # allowed: daily|weekly|monthly
    QC_variables = [v.strip() for v in params.get("QC_variables", []) if v.strip()]
    columns_breakdown = [
        v.strip() for v in params.get("columns_breakdown", []) if v.strip()
    ]
    raw_metrics = [m.strip() for m in params.get("metrics", []) if m.strip()]
    metric_var = params.get("metric_var", "metric").strip()
    value_var = params.get("value_var", "value").strip()
    cost_var = params.get("cost_var", "Cost").strip()

    priority_map = {
        "impressions": 1,
        "sent": 2,
        "sents": 2,
        "send": 2,
        "sends": 2,
        "delivered": 3,
        "opens": 4,
        "open": 4,
        "clicks": 5,
        "cost": 6,
    }

    _col_map = {c.strip().lower(): c for c in df.columns}

    incomplete_quarters_all = []

    def check_incomplete_quarter(
        pivot_week, pivot_q, date_var, date_grain, sheet_name, ws, q_startcol
    ):
        if date_grain != "weekly" or pivot_week.empty:
            return

        max_date = pd.to_datetime(pivot_week[date_var].max())
        weekday = max_date.strftime("%A")
        quarter_end_date = (max_date + QuarterEnd(0)).normalize()
        while quarter_end_date.strftime("%A") != weekday:
            quarter_end_date -= pd.Timedelta(days=1)

        if max_date != quarter_end_date:
            quarter_str = max_date.to_period("Q").strftime("%Y Q%q")
            # Mark incomplete quarter in pivot_q
            pivot_q["Quarter"] = pivot_q["Quarter"].apply(
                lambda q: q + "*" if q == quarter_str else q
            )
            # Add footnote in Excel
            ws.write(len(pivot_q) + 1, q_startcol, "* Incomplete Quarter")
            # Save details for summary sheet
            incomplete_quarters_all.append(
                {
                    "Sheet Name": sheet_name,
                    "Quarter": quarter_str,
                    "Max Date": max_date.strftime("%Y-%m-%d"),
                    "Weekday": weekday,
                    "Quarter End": quarter_end_date.strftime("%Y-%m-%d"),
                    "Status": "Incomplete",
                }
            )

    def _resolve(col_name: str) -> str:
        """Resolve a requested column name to the actual df column name (case/space insensitive)."""
        return _col_map.get(col_name.strip().lower(), col_name)

    metric_var_res = _resolve(metric_var)
    value_var_res = _resolve(value_var)
    cost_var_res = _resolve(cost_var)

    has_long = (metric_var_res in df.columns) and (value_var_res in df.columns)

    if not has_long and raw_metrics:
        wanted = {m.lower() for m in raw_metrics}
        Metrics = sorted(
            [col for col in df.columns if col.strip().lower() in wanted],
            key=lambda x: priority_map.get(x.strip().lower(), 99),
        )
        if not Metrics:
            raise KeyError(
                f"No metric columns from params.metrics={raw_metrics} were found in the input file."
            )
    elif has_long:
        df[metric_var_res] = df[metric_var_res].astype(str).str.strip()
        Metrics = sorted(
            df[metric_var_res].dropna().unique().tolist(),
            key=lambda x: priority_map.get(str(x).strip().lower(), 99),
        )
    else:
        raise KeyError(
            "Could not detect input format: "
            f"Expected long columns '{metric_var}'/'{value_var}' (found? "
            f"{metric_var_res in df.columns}/{value_var_res in df.columns}) or "
            f"a non-empty 'metrics' list for wide input."
        )

    agg_dict = {m: "sum" for m in Metrics} if not has_long else None

    metric_var = metric_var_res
    value_var = value_var_res
    cost_var = cost_var_res

    df[date_var] = pd.to_datetime(df[date_var], errors="coerce")

    if QC_variables:
        QC_dict = {
            f"QC{i}": sorted(df[v].dropna().unique())
            for i, v in enumerate(QC_variables)
        }
        QC_combinations = list(
            itertools.product(*[QC_dict[f"QC{i}"] for i in range(len(QC_variables))])
        )
    else:
        QC_combinations = [()]

    out_path = Path(
        params.get("output_path")
        or csv_path.with_name(f"EDA_general_{pd.Timestamp.now():%H%M%S}.xlsx")
    )

    with pd.ExcelWriter(out_path, engine="xlsxwriter") as writer:
        workbook = writer.book
        workbook.nan_inf_to_errors = True

        # number_fmt   = workbook.add_format({'num_format': '#,##0', 'border': 1})
        number_fmt = workbook.add_format(
            {"num_format": '_(* #,##0_);_(* (#,##0);_(* "-"??_);_(@_)', "border": 1}
        )
        currency_fmt = workbook.add_format(
            {"num_format": '_($* #,##0_);_($* (#,##0);_($* "-"??_);_(@_)', "border": 1}
        )
        header_fmt = workbook.add_format(
            {
                "bold": True,
                "bg_color": Tab_Color_Palette,
                "font_color": "white",
                "align": "center",
                "valign": "vcenter",
                "text_wrap": True,
                "border": 1,
            }
        )
        # date_fmt     = workbook.add_format({'num_format': 'yyyy-mm-dd', 'border': 1})
        date_fmt = workbook.add_format({"num_format": "m/d/yyyy", "border": 1})
        border_fmt = workbook.add_format({"border": 1})

        for idx, combo in enumerate(QC_combinations, 1):
            df_f = df.copy()

            if QC_variables:
                for i, val in enumerate(combo):
                    df_f = df_f[df_f[QC_variables[i]] == val]

            if df_f.empty:
                continue
            valid_columns_breakdown = [
                col
                for col in columns_breakdown
                if col in df_f.columns
                and df_f[col].dropna().astype(str).str.strip().ne("").any()
            ]

            group_cols = [date_var] + QC_variables + valid_columns_breakdown
            if has_long:
                df_week_long = df_f.groupby(group_cols + [metric_var], as_index=False)[
                    value_var
                ].sum()
                long = df_week_long.rename(
                    columns={metric_var: "metric", value_var: "value"}
                )
            else:
                df_week = df_f.groupby(group_cols, as_index=False).agg(agg_dict)
                df_week = df_week[group_cols + Metrics]
                id_vars = [date_var] + valid_columns_breakdown
                long = df_week.melt(
                    id_vars=id_vars,
                    value_vars=Metrics,
                    var_name="metric",
                    value_name="value",
                )

            metric_order = {m: i + 1 for i, m in enumerate(Metrics)}
            long["metric_order"] = long["metric"].map(metric_order)
            long["combined"] = (
                long[valid_columns_breakdown].agg(" ".join, axis=1)
                + " "
                + long["metric"]
                if valid_columns_breakdown
                else long["metric"]
            )
            long = long[[date_var, "combined", "value"]]

            pivot_week = (
                long.groupby([date_var, "combined"], as_index=False)["value"]
                .sum()
                .pivot(index=date_var, columns="combined", values="value")
                .fillna(0)
                .reset_index()
            )
            sorted_cols = [date_var] + sorted(
                pivot_week.columns[1:],
                key=lambda c: (
                    1 if c.lower().startswith("total") else 0,
                    priority_map.get(c.rsplit(" ", 1)[-1].lower(), 99),
                ),
            )
            pivot_week = pivot_week[sorted_cols]

            total_cols = {}
            for metric in Metrics:
                metric_cols = [
                    c
                    for c in pivot_week.columns
                    if c.rsplit(" ", 1)[-1].lower() == metric.lower()
                ]
                if metric_cols:
                    total_name = f"Total {metric.capitalize()}"
                    pivot_week[total_name] = pivot_week[metric_cols].sum(axis=1)
                    total_cols[metric] = total_name

            final_week_cols = (
                [date_var]
                + [
                    c
                    for c in pivot_week.columns
                    if c not in total_cols.values() and c != date_var
                ]
                + list(total_cols.values())
            )
            pivot_week = pivot_week[final_week_cols]

            for col in list(pivot_week.columns):
                if col.lower().startswith("total "):
                    base_metric = col.replace("Total ", "").strip()
                    breakdown_cols = [
                        c
                        for c in pivot_week.columns
                        if base_metric in c and not c.lower().startswith("total")
                    ]
                    if len(breakdown_cols) <= 1:
                        pivot_week = pivot_week.drop(columns=[col])

            if not pivot_week.empty:
                pivot_week[date_var] = pd.to_datetime(
                    pivot_week[date_var], errors="coerce"
                )
                dmin = pd.to_datetime(pivot_week[date_var].min())
                dmax = pd.to_datetime(pivot_week[date_var].max())

                if date_grain == "daily":
                    idx = pd.date_range(dmin.normalize(), dmax.normalize(), freq="D")

                elif date_grain == "weekly":
                    anchor = dmin.strftime("%a").upper()[:3]
                    freq = f"W-{anchor}"
                    idx = pd.date_range(dmin, dmax, freq=freq)

                elif date_grain == "monthly":
                    dmin_ms = dmin.to_period("M").to_timestamp("MS")
                    dmax_ms = dmax.to_period("M").to_timestamp("MS")
                    idx = pd.date_range(dmin_ms, dmax_ms, freq="MS")

                else:
                    idx = pd.date_range(
                        dmin.normalize(), dmax.normalize(), freq=f"W-{anchor}"
                    )
                pivot_week = (
                    pivot_week.set_index(date_var)
                    .reindex(idx)
                    .fillna(0)
                    .rename_axis(date_var)
                    .reset_index()
                )
            df_f = df_f[df_f[date_var].notna()]
            df_f["Quarter"] = (
                df_f[date_var].dt.to_period("Q").astype(str).str.replace("Q", " Q")
            )
            q_group = ["Quarter"] + QC_variables + valid_columns_breakdown
            if has_long:
                df_q_long = df_f.groupby(q_group + [metric_var], as_index=False)[
                    value_var
                ].sum()
                q_long = df_q_long.rename(
                    columns={metric_var: "metric", value_var: "value"}
                )
            else:
                df_q = df_f.groupby(q_group, as_index=False).agg(agg_dict)
                q_id = ["Quarter"] + valid_columns_breakdown
                q_long = df_q.melt(
                    id_vars=q_id,
                    value_vars=Metrics,
                    var_name="metric",
                    value_name="value",
                )
            q_long["combined"] = (
                q_long[valid_columns_breakdown].agg(" ".join, axis=1)
                + " "
                + q_long["metric"]
                if valid_columns_breakdown
                else q_long["metric"]
            )
            q_long = q_long[["Quarter", "combined", "value"]]
            pivot_q = (
                q_long.groupby(["Quarter", "combined"], as_index=False)["value"]
                .sum()
                .pivot(index="Quarter", columns="combined", values="value")
                .fillna(0)
                .reset_index()
            )
            q_sorted = ["Quarter"] + sorted(
                pivot_q.columns[1:],
                key=lambda c: (
                    1 if c.lower().startswith("total") else 0,
                    priority_map.get(c.rsplit(" ", 1)[-1].lower(), 99),
                ),
            )

            pivot_q = pivot_q[q_sorted]

            for metric in Metrics:
                metric_cols = [
                    c
                    for c in pivot_q.columns
                    if c.rsplit(" ", 1)[-1].lower() == metric.lower()
                ]
                if metric_cols:
                    total_name = f"Total {metric.capitalize()}"
                    pivot_q[total_name] = pivot_q[metric_cols].sum(axis=1)

            non_total_cols = [c for c in pivot_q.columns if not c.startswith("Total ")]
            total_cols_list = [c for c in pivot_q.columns if c.startswith("Total ")]
            pivot_q = pivot_q[non_total_cols + total_cols_list]

            for col in list(pivot_q.columns):
                if col.lower().startswith("total "):
                    base_metric = col.replace("Total ", "").strip()
                    breakdown_cols = [
                        c
                        for c in pivot_q.columns
                        if base_metric in c and not c.lower().startswith("total")
                    ]
                    if len(breakdown_cols) <= 1:
                        pivot_q = pivot_q.drop(columns=[col])

            full_sheet = "_".join(map(str, combo)) if QC_variables else "General"
            short_sheet = short_name(full_sheet)[:31]

            non_date_cols = [col for col in pivot_week.columns if col != date_var]
            if len(non_date_cols) == 2:
                total_col = next(
                    (col for col in non_date_cols if col.lower().startswith("total")),
                    None,
                )
                if total_col:
                    pivot_week = pivot_week.drop(columns=[total_col])

            pivot_week.to_excel(writer, sheet_name=short_sheet, index=False)
            ws = writer.sheets[short_sheet]
            q_startcol = len(pivot_week.columns) + 1
            pivot_q.to_excel(
                writer, sheet_name=short_sheet, index=False, startcol=q_startcol
            )
            check_incomplete_quarter(
                pivot_week, pivot_q, date_var, date_grain, short_sheet, ws, q_startcol
            )

            def fmt_header(df_p, sc=0):
                for c, val in enumerate(df_p.columns):
                    if val == date_var:
                        display = "Date"
                    else:
                        display = pretty_col(val)

                    ws.write(0, sc + c, display, header_fmt)

            def fmt_cols(df_p, sc=0, is_q=False):
                for c, col in enumerate(df_p.columns):
                    ci = sc + c
                    vals = df_p[col].apply(
                        lambda x: f"{x:,}" if isinstance(x, (int, float)) else str(x)
                    )
                    w = max(vals.astype(str).map(len).max(), len(col)) + 2
                    ws.set_column(ci, ci, w)
                    fmt = number_fmt
                    if date_var == col:
                        the_fmt = date_fmt
                        is_date_col = True
                    elif cost_var and cost_var.lower() in col.lower():
                        the_fmt = currency_fmt
                        is_date_col = False
                    elif is_q and "quarter" in col.lower():
                        the_fmt = border_fmt
                        is_date_col = False
                    else:
                        the_fmt = number_fmt
                        is_date_col = False

                    for r in range(1, len(df_p) + 1):
                        val = df_p.iloc[r - 1, c]
                        if is_date_col:
                            py_dt = pd.to_datetime(val, errors="coerce")
                            if pd.notna(py_dt):
                                ws.write_datetime(r, ci, py_dt.to_pydatetime(), the_fmt)
                            else:
                                ws.write(r, ci, "", the_fmt)
                        else:
                            ws.write(r, ci, val, the_fmt)

            fmt_header(pivot_week)
            fmt_cols(pivot_week)
            fmt_header(pivot_q, q_startcol)
            fmt_cols(pivot_q, q_startcol, is_q=True)

            ws.freeze_panes(1, 1)
            ws.set_zoom(100)

            num_rows = len(pivot_week)
            total_week_cols = len(pivot_week.columns) - 1
            for m_idx, metric in enumerate(Metrics):
                mcols = [
                    c
                    for c in pivot_week.columns
                    if (metric.lower() in c.lower())
                    and not c.strip().lower().startswith("total ")
                ]
                if not mcols:
                    continue

                if len(mcols) <= 2:
                    chart = workbook.add_chart({"type": "line"})
                    chart.set_title(
                        {
                            "name": pretty_col(metric)
                            if pretty_col(short_sheet) == pretty_col(metric)
                            else f"{pretty_col(short_sheet)} {pretty_col(metric)}",
                            "name_font": {"color": "#595959"},
                        }
                    )
                    chart.set_x_axis(
                        {  #'name': 'Date',
                            "date_axis": True,
                            "num_format": "mmm `yy",
                            "major_gridlines": {"visible": False},
                            "minor_gridlines": {"visible": False},
                            #'num_font': {'rotation': 45}
                        }
                    )
                    chart.set_y_axis(
                        {  #'name': pretty_col(metric),
                            "number_format": "#,##0",
                            "line": {"none": True},
                            "major_gridlines": {"visible": False},
                            "minor_gridlines": {"visible": False},
                        }
                    )

                    for k, col in enumerate(mcols):
                        clean = pretty_col(col).replace("\n", " ")
                        loc = pivot_week.columns.get_loc(col)
                        chart.add_series(
                            {
                                "name": clean,
                                "categories": f"='{short_sheet}'!$A$2:$A${num_rows + 1}",
                                "values": f"='{short_sheet}'!${df_columns[loc]}$2:${df_columns[loc]}${num_rows + 1}",
                                "line": {
                                    "color": Color_Palette[k % len(Color_Palette)],
                                    "width": 2,
                                },
                            }
                        )

                else:
                    chart = workbook.add_chart({"type": "area", "subtype": "stacked"})
                    chart.set_title(
                        {
                            "name": pretty_col(metric)
                            if pretty_col(short_sheet) == pretty_col(metric)
                            else f"{pretty_col(short_sheet)} {pretty_col(metric)}",
                            "name_font": {"color": "#595959"},
                        }
                    )
                    chart.set_x_axis(
                        {  #'name': 'Date',
                            "date_axis": True,
                            "num_format": "mmm `yy",
                            "major_gridlines": {"visible": False},
                            "minor_gridlines": {"visible": False},
                            #'num_font': {'rotation': 45}
                        }
                    )
                    chart.set_y_axis(
                        {  #'name': pretty_col(metric),
                            "number_format": "#,##0",
                            "line": {"none": True},
                            "major_gridlines": {"visible": False},
                            "minor_gridlines": {"visible": False},
                        }
                    )

                    for k, col in enumerate(mcols):
                        clean = pretty_col(col).replace("\n", " ")
                        loc = pivot_week.columns.get_loc(col)
                        chart.add_series(
                            {
                                "name": clean,
                                "categories": f"='{short_sheet}'!$A$2:$A${num_rows + 1}",
                                "values": f"='{short_sheet}'!${df_columns[loc]}$2:${df_columns[loc]}${num_rows + 1}",
                                "fill": {
                                    "color": Color_Palette[k % len(Color_Palette)]
                                },
                                "line": {
                                    "color": Color_Palette[k % len(Color_Palette)],
                                    "width": 2,
                                },
                                "border": {"none": True},
                            }
                        )

                if len(mcols) == 1:
                    chart.set_legend({"none": True})
                else:
                    chart.set_legend({"position": "bottom"})

                chart.set_size({"width": 850, "height": 450})

                # chart_row = len(pivot_q) + 3
                # chart_spacing_cols = 15
                # if len(Metrics) == 1:
                #     chart_col = 25
                # else:
                #     chart_col = q_startcol + m_idx * chart_spacing_cols
                # ws.insert_chart(chart_row, chart_col, chart)

                base_row = len(pivot_q) + 3
                cols_per_chart = 25

                for m_idx, metric in enumerate(Metrics):
                    chart_row = base_row
                    chart_col = q_startcol + (m_idx * cols_per_chart)
                    ws.insert_chart(chart_row, chart_col, chart)

                ############### Dual axis chart portion###############
            if cost_var and cost_var.strip():  # Only proceed if cost_var is provided
                total_cols = [
                    c for c in pivot_week.columns if c.lower().startswith("total ")
                ]
                cost_col = None
                if total_cols:
                    base_names = [c.replace("Total ", "").strip() for c in total_cols]
                    for idx, base in enumerate(base_names):
                        if base.lower() == cost_var.lower():
                            cost_col = total_cols[idx]
                            break
                if not cost_col:
                    cost_candidates = [
                        c for c in pivot_week.columns if cost_var.lower() in c.lower()
                    ]
                    if cost_candidates:
                        cost_col = cost_candidates[0]
                for metric in Metrics:
                    if metric.lower() == cost_var.lower():
                        continue
                    if total_cols:
                        primary_col = (
                            f"Total {metric.capitalize()}"
                            if f"Total {metric.capitalize()}" in pivot_week.columns
                            else None
                        )
                    else:
                        primary_col = metric if metric in pivot_week.columns else None
                    if not primary_col or not cost_col:
                        continue

                    chart = workbook.add_chart({"type": "line"})
                    chart.set_title(
                        {
                            "name": f"{pretty_col(primary_col)} vs {pretty_col(cost_col)}",
                            "name_font": {"color": "#595959"},
                        }
                    )
                    chart.set_x_axis(
                        {
                            "date_axis": True,
                            "num_format": "mmm `yy",
                            "major_gridlines": {"visible": False},
                            "minor_gridlines": {"visible": False},
                            #'num_font': {'rotation': 45}
                        }
                    )
                    chart.set_y_axis(
                        {  #'name': pretty_col(primary_col),
                            "number_format": "#,##0",
                            "line": {"none": True},
                            "major_gridlines": {"visible": False},
                            "minor_gridlines": {"visible": False},
                        }
                    )

                    loc_primary = pivot_week.columns.get_loc(primary_col)
                    chart.add_series(
                        {
                            "name": pretty_col(primary_col),
                            "categories": f"='{short_sheet}'!$A$2:$A${num_rows + 1}",
                            "values": f"='{short_sheet}'!${df_columns[loc_primary]}$2:${df_columns[loc_primary]}${num_rows + 1}",
                            "line": {"color": Color_Palette[0], "width": 2},
                            "y_axis": True,
                        }
                    )

                    chart.set_y2_axis(
                        {  #'name': pretty_col(primary_col),
                            "number_format": "#,##0",
                            "line": {"none": True},
                            "major_gridlines": {"visible": False},
                            "minor_gridlines": {"visible": False},
                        }
                    )
                    loc_cost = pivot_week.columns.get_loc(cost_col)
                    chart.add_series(
                        {
                            "name": pretty_col(cost_col),
                            "categories": f"='{short_sheet}'!$A$2:$A${num_rows + 1}",
                            "values": f"='{short_sheet}'!${df_columns[loc_cost]}$2:${df_columns[loc_cost]}${num_rows + 1}",
                            "line": {"color": Color_Palette[1], "width": 2},
                            "y2_axis": True,
                        }
                    )

                    chart.set_legend({"position": "bottom"})
                    chart.set_size({"width": 850, "height": 450})

                    chart_spacing_cols = 20
                    base_row = len(pivot_q) + 27
                    start_col = len(pivot_week.columns) + 1
                    chart_row = base_row
                    chart_col = start_col
                    ws.insert_chart(chart_row, chart_col, chart)

                ############### Dual axis chart ###############

        df_summary = df.copy()
        df_summary[date_var] = pd.to_datetime(df_summary[date_var], errors="coerce")
        df_summary["Year"] = df_summary[date_var].dt.year
        df_summary["Month"] = df_summary[date_var].dt.month

        group_cols_summary = ["Year", "Month"] + QC_variables + valid_columns_breakdown

        if has_long:
            df_grouped_long = df_summary.groupby(
                group_cols_summary + [metric_var], as_index=False
            )[value_var].sum()
            df_grouped = df_grouped_long.pivot_table(
                index=group_cols_summary,
                columns=metric_var,
                values=value_var,
                aggfunc="sum",
                fill_value=0,
            ).reset_index()
            metric_cols = [c for c in df_grouped.columns if c not in group_cols_summary]
        else:
            metric_cols = Metrics
            df_grouped = df_summary.groupby(group_cols_summary, as_index=False)[
                metric_cols
            ].sum()
        sort_cols = ["Year", "Month"] + QC_variables + valid_columns_breakdown
        df_grouped = df_grouped.sort_values(sort_cols).reset_index(drop=True)
        # if incomplete_quarters_all:
        #     pd.DataFrame(incomplete_quarters_all).to_excel(writer, sheet_name="Incomplete Quarters", index=False)
        df_grouped.to_excel(writer, sheet_name="Summary", index=False)
        ws_summary = writer.sheets["Summary"]
        cost_format = workbook.add_format(
            {"num_format": '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)'}
        )
        # int_format  = workbook.add_format({'num_format': '#,##0'})
        int_format = workbook.add_format(
            {"num_format": '_(* #,##0_);_(* (#,##0);_(* "-"??_);_(@_)'}
        )
        bold_format = workbook.add_format({"bold": True})

        col_indices = {col: idx for idx, col in enumerate(df_grouped.columns)}
        ws_summary.set_column(col_indices["Year"], col_indices["Year"], 12, bold_format)
        ws_summary.set_column(
            col_indices["Month"], col_indices["Month"], 12, bold_format
        )

        for col in metric_cols:
            if col not in col_indices:
                continue
            idx = col_indices[col]
            if "spend" in col.lower() or "cost" in col.lower():
                ws_summary.set_column(idx, idx, 15, cost_format)
            else:
                ws_summary.set_column(idx, idx, 15, int_format)

        for c, col in enumerate(df_grouped.columns):
            ws_summary.write(
                0, c, "Date" if col == date_var else pretty_col(col), header_fmt
            )

    if out_path.exists():
        print(f"EDA saved: {out_path}")
    #    if os.name == "nt":
    #       os.startfile(str(out_path))

    return str(out_path)
