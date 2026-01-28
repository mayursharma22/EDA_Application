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
- **`src/eda_generation/eda_ppt_generation.py`** — the PPT writer/formatter that builds summary sheets and charts.

---

# 🧭 Usage

## 1) Data Preparation (CSV/XLSX → CSV)
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
2. Upload Wide.Semi-Wide format datasets.
3. Configure date/channel/columns.
4. Melt/group/export.

---

## 2) EDA Generation (CSV → Excel & PPT)
**What it does**
- Upload long format CSV and preview.
- Configure date, granularity, metrics, breakdowns.
- Customize visual colors.
- Generate Excel & PPT Deck with pivots, charts, summary sheets.

**How to use**
1. Choose **EDA Generation**.
2. Choose parameters.
3. Export Excel & PPT Deck.

---

# 🧩 Design Notes
- **Data Preparation** accepts CSV/XLSX in Wide/Semi-Wide format; **EDA Generation**  expects CSV in long format.
- Sheet names ≤ 31 chars.

---

# 🛠 Troubleshooting
- Validate color hex codes.
- Clean numeric columns before converting.

---

# 🌐 App URL
- https://eda-application-3365883928105847.aws.databricksapps.com/

---

# 📌 Run Application in Local Machine

## 🖥️ Guide to Setup Application in Local Machine (Windows)
Below is the complete, step‑by‑step one time process to setup application in local machine for windows machine

---

### ✅ Prerequisites
- Windows OS
- PowerShell (Admin)
- Git installed
- Internet connection

---

### 🔧 Step 1: Install UV (Python Environment Manager)
UV is a fast environment & package manager. To Download UV in your Machine Open PowerShell in your Windows Machine and copy the following command over there

```
powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"
```
![alt text](image.png)

---

### 📥 Step 2: Clone the Repository
1. Visit GitHub repo: To Clone the EDA_Application source code from GitHub to your local machine open below link in your browser
```
https://github.com/mayursharma22/EDA_Application
```

2. Copy HTTPS clone URL.

3. In terminal: Open cmd terminal your desired directory where you want to clone the EDA_Application Repository and run below command.
```
git clone https://github.com/mayursharma22/EDA_Application.git
```
![alt text](image-1.png)

---

### 📂 Step 3: Set Up Environment
1. Navigate to cloned repository folder: Go to the location where you cloned the repository
![alt text](image-2.png)

2. Install the required dependencies: Open Terminal (cmd/GitBash) and install dependencies using below command line. 
```
uv sync -n
```
![alt text](image-3.png)

---

## ▶️ Run the Application 
1. Run Streamlit using UV: Once setup Application in Local Machine, you have to run only below command in your Terminal from the folder where you cloned the repository using cmd/Git Bash
```
uv run streamlit run app.py --server.address="localhost"
```
![alt text](image-4.png)

---

2. 🎉 Expected Result
Your app launches at: http://localhost:8000 If not, open the URL manually.

---
