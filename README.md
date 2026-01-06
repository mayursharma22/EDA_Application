# **EDA_Application**

A Streamlit-powered toolkit that helps user **prepare raw marketing/analytics data** (single or multi-file) and **generate a formatted Excel EDA workbook** with pivots, totals, and publication-ready charts. The app ships with two tools accessible from the sidebar:

- **Data Preparation** — interactively clean, harmonize, *melt*, and group datasets; export a single CSV or a ZIP split by a chosen dimension.
- **EDA Generation** — upload a CSV and produce a polished Excel workbook with weekly/monthly/daily pivots, totals, quarterly views, and color‑customized charts.

---

## 🚀 Quick Start

When the app opens in your browser, use the sidebar to switch between **Data Preparation** and **EDA Generation**.

---

## 📁 Repository Structure

```
EDA_Application/
├─ app.py
├─ src/
│  ├─ data_preperation/
│  │  └─ eda_data_processing.py
│  └─ eda_generation/
│     ├─ eda_excel_app.py
│     └─ eda_excel_generation.py
```

- **`app.py`** — Streamlit entry point and navigation between the two tools.
- **`src/data_preperation/eda_data_processing.py`** — the **Data Preparation** tool (single-file and multi-file flows).
- **`src/eda_generation/eda_excel_app.py`** — the **EDA Generation** Streamlit UI that gathers parameters and triggers Excel creation.
- **`src/eda_generation/eda_excel_generation.py`** — the Excel writer/formatter that builds sheets, totals, and charts.

---

## 🧭 Usage

### 1) Data Preparation (CSV/XLSX)

**What it does**

- Upload one or multiple files (**CSV** or **Excel**) and preview them.
- Pick/rename a **date column** with robust parsing.
- Specify a **channel** column (use existing, rename, or create constant).
- Rename other columns & add new fields.
- Detect numeric-like columns safely before converting.
- Melt selected metric columns to long format.
- Group by chosen dimensions and sum `Values`.
- Multi-file harmonization and breakdown export.

**How to use**

1. In the sidebar, choose **Data Preparation**. Upload one or more CSV/XLSX files.
2. Configure date, channel, and other columns.
3. Apply melt, group, and export CSV or ZIP.

---

### 2) EDA Generation (CSV → Excel)

**What it does**

- Upload a CSV, preview, and select parameters.
- Customize colors for charts and headers.
- Generate Excel with pivots, charts, and summary sheets.

**How to use**

1. In the sidebar, choose **EDA Generation**.
2. Configure date, granularity, metrics, breakdowns.
3. Customize colors and export Excel workbook.

---

## 🧱 Design Notes & Constraints

- File types: Data Preparation supports CSV/XLSX; EDA Generation expects CSV.
- Sheet names shortened to ≤31 chars.
- Currency formatting applied to cost/spend columns.

---

## 🧪 Tips

- Set Week Start Day for weekly alignment.
- Provide HEX colors for charts and headers.

---

## 🛠 Troubleshooting

- Validate colors before proceeding.
- Ensure numeric columns are cleaned.

---

## 🧪 App URL

- https://eda-application-3365883928105847.aws.databricksapps.com/

---

## 🧪 Local Run

To run the app locally, after installing packages from requirements.txt
- streamlit run ./app.py --server.address="localhost"
