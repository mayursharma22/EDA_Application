# **EDA_Application**

A Streamlit-powered toolkit that helps user **prepare raw marketing/analytics data** (single or multi-file) and **generate a formatted Excel EDA workbook** with pivots, totals, and publication-ready charts. The app ships with two tools accessible from the sidebar:

- **Data Preparation** — interactively clean, harmonize, *melt*, and group datasets; export a single CSV or a ZIP split by a chosen dimension.
- **EDA Generation** — upload a CSV and produce a polished Excel workbook with weekly/monthly/daily pivots, totals, quarterly views, and color‑customized charts.

---

# 🚀 Quick Start
When the app opens in your browser, use the sidebar to switch between **Data Preparation** and **EDA Generation**.

---

# 📁 Repository Structure
```
EDA_Application/
├── app.py
├── src/
│   ├── data_preperation/
│   │   └── eda_data_processing.py
│   └── eda_generation/
│       ├── eda_excel_app.py
│       └── eda_excel_generation.py
        └── eda_ppt_generation.py
```

- **`app.py`** — Streamlit entry point and navigation between the two tools.
- **`src/data_preperation/eda_data_processing.py`** — the **Data Preparation** tool (single-file and multi-file flows).
- **`src/eda_generation/eda_excel_app.py`** — the **EDA Generation** Streamlit UI that gathers parameters and triggers Excel creation.
- **`src/eda_generation/eda_excel_generation.py`** — the Excel writer/formatter that builds sheets, totals, and charts.

---

# 🧭 Usage
## 1) Data Preparation (CSV/XLSX)
**What it does**
- Upload one or more files (CSV or Excel) and preview them.
- Pick/rename a **date column** with robust parsing.
- Create/use a **channel column**.
- Rename columns, add custom fields.
- Melt selected metric columns.
- Group by chosen dimensions.
- Export harmonized CSV or ZIP breakdown.

**How to use**
1. In sidebar, choose **Data Preparation**.
2. Upload datasets.
3. Configure date/channel/columns.
4. Melt/group/export.

---

## 2) EDA Generation (CSV → Excel)
**What it does**
- Upload CSV and preview.
- Configure date, granularity, metrics, breakdowns.
- Customize visual colors.
- Generate Excel with pivots, charts, summary sheets.

**How to use**
1. Choose **EDA Generation**.
2. Choose parameters.
3. Export Excel.

---

# 🧩 Design Notes
- Data Prep accepts CSV/XLSX; EDA expects CSV.
- Sheet names ≤ 31 chars.
- Currency formatting for cost/spend.

---

# 🛠 Troubleshooting
- Validate color hex codes.
- Clean numeric columns before converting.

---

# 🌐 App URL
- https://eda-application-3365883928105847.aws.databricksapps.com/

---

# 🖥️ Local Run (Updated with Full Guide)
Below is the complete, step‑by‑step local run process combining README and the Word setup guide.

---

# 📌 Local Setup Guide (Windows)

## ✅ Prerequisites
- Windows OS
- PowerShell (Admin)
- Git installed
- Internet connection

---

## 🔧 Step 1: Install UV (Python Environment Manager)
UV is a fast environment & package manager.

### 1. Open UV installation page:
https://docs.astral.sh/uv/getting-started/installation/

### 2. Run installation command in PowerShell:
```
powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"
```

---

## 📥 Step 2: Clone the Repository
1. Visit GitHub repo:
https://github.com/mayursharma22/EDA_Application

2. Copy HTTPS clone URL.

3. In terminal:
```
git clone https://github.com/mayursharma22/EDA_Application.git
```

---

## 📂 Step 3: Set Up Environment
Navigate to cloned folder and install dependencies:
```
uv sync -n
```

---

## ▶️ Step 4: Run the Application
Run Streamlit using UV:
```
uv run streamlit run app.py --server.address="localhost"
```

---

## 🎉 Expected Result
Your app launches at:
http://localhost:8000
If not, open the URL manually.

---
