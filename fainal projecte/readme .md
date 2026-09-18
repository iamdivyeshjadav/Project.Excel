<div align="center">

# 📊 Sales Performance Dashboard — Excel Project
### *Interactive Excel Workbook: Data Cleaning, Pivot Analysis, Advanced Formulas & Dashboard*

[![Excel](https://img.shields.io/badge/Excel-2016%2B-217346?style=for-the-badge&logo=microsoftexcel&logoColor=white)](https://www.microsoft.com/excel)
[![Formulas](https://img.shields.io/badge/Formulas-XLOOKUP%20%7C%20INDEX--MATCH%20%7C%20SUMIFS-FF6F00?style=for-the-badge&logo=microsoftexcel&logoColor=white)](https://www.microsoft.com/excel)
[![Dashboard](https://img.shields.io/badge/Dashboard-Interactive-4CAF50?style=for-the-badge&logo=googlesheets&logoColor=white)](https://www.microsoft.com/excel)
[![Dataset](https://img.shields.io/badge/Dataset-UCI%20Online%20Retail-9C27B0?style=for-the-badge&logo=databricks&logoColor=white)](https://archive.ics.uci.edu/)

<br/>

> *"Raw data tells a story — pivots, formulas and dashboards just help you read it."*

</div>

---

## 📋 Table of Contents

- [📌 Overview](#-overview)
- [🎯 Objective](#-objective)
- [🏗️ Workbook Structure](#️-workbook-structure)
- [🔄 Project Workflow](#-project-workflow)
- [🧹 Data Cleaning](#-data-cleaning)
- [📊 Pivot-Style Summary Tables](#-pivot-style-summary-tables)
- [🧮 Advanced Formulas](#-advanced-formulas)
- [📈 Interactive Dashboard](#-interactive-dashboard)
- [🗂️ Extra Pivot Tables](#️-extra-pivot-tables)
- [🛠️ Tech Stack](#️-tech-stack)
- [🏆 Marks Mapping](#-marks-mapping)
- [📄 License](#-license)
- [👤 Author](#-author)
- [🙏 Acknowledgements](#-acknowledgements)

---

## 📌 Overview

The **Sales Performance Dashboard** is an Excel-based data analysis project built on the
**UCI "Online Retail"** dataset. It takes raw UK online-retail transactions through
**cleaning → pivot summaries → advanced formulas → an interactive dashboard**, all inside a
single `.xlsx` workbook — no external tools or add-ins required.

This project is designed to:
- Practice real-world spreadsheet data cleaning (duplicates, blanks, standardization)
- Build formula-driven pivot-style summaries that refresh automatically
- Demonstrate advanced Excel formulas (XLOOKUP, INDEX-MATCH, SUMIFS, RANK, nested IF)
- Present findings through KPI cards, charts, and a dropdown-driven interactive dashboard

---

## 🎯 Objective

> **Goal:** Turn a raw retail transactions export into a clean, analyzable dataset and present
> the results through pivot tables, formulas, and a one-page interactive dashboard.

| 📂 Component | 📄 Type | 🔍 Description |
|---|---|---|
| Raw_Data | Sheet | Original imported dataset, intentionally messy |
| Clean_Data | Sheet | Deduplicated, filled, standardized, with calculated fields |
| Pivot_Summary | Sheet | Formula-based pivot-style tables (country, month, segment) |
| Advanced_Formulas | Sheet | XLOOKUP, INDEX-MATCH, IF, SUMIFS, COUNTIFS, RANK demos |
| pivot_table | Sheet | Native-style pivot breakdowns by country and month |
| Dashboard | Sheet | KPI cards, charts, and country/year selectors |
| Documentation | Sheet | Project notes and submission guide |

---

## 🏗️ Workbook Structure

```
📦 fainal_project.xlsx
│
├── 📄 Raw_Data           ← Original imported dataset (with blanks/duplicate)
├── 📄 Clean_Data          ← Cleaned + enriched dataset (Revenue, Month, Year, Segment)
├── 📄 Pivot_Summary        ← Formula-based pivot-style summary tables
├── 📄 Advanced_Formulas     ← XLOOKUP / INDEX-MATCH / SUMIFS / RANK demonstrations
├── 📄 pivot_table          ← Country & month pivot breakdowns
├── 📄 Dashboard            ← KPI cards + charts + interactive selectors
└── 📄 Documentation        ← Submission guide and dataset notes
```

---

## 🔄 Project Workflow

```
Raw_Data (messy import)
      │
      ▼
┌───────────────────────────────┐
│   Clean_Data                  │  ← Remove duplicate, fill blanks,
│   (cleaning + enrichment)     │    standardize status, add Revenue/Month/Year/Segment
└───────────────┬───────────────┘
                │
        ┌───────┴────────┐
        ▼                ▼
┌───────────────┐   ┌──────────────────┐
│ Pivot_Summary  │   │ Advanced_Formulas│
│ (country/month/│   │ (lookup, logic,  │
│  segment)      │   │  ranking demos)  │
└───────┬────────┘   └────────┬─────────┘
        │                     │
        └──────────┬──────────┘
                    ▼
          ┌───────────────────┐
          │     Dashboard      │  ← KPI cards, charts,
          │ (interactive view) │    Country & Year selectors
          └───────────────────┘
```

---

## 🧹 Data Cleaning

`Clean_Data` is built from `Raw_Data` and:
- Removes the intentionally-included duplicate row
- Fills missing `Quantity` / `CustomerID` values
- Standardizes cancellation status into a consistent `OrderStatus` field
- Adds calculated columns: `Revenue`, `Month`, `Year`, `CustomerSegment`

**Columns:**

| Field | Description |
|---|---|
| `InvoiceNo` | Invoice number (a `C` prefix marks a cancellation) |
| `StockCode` | Product code |
| `Description` | Product name |
| `Quantity` | Units purchased |
| `InvoiceDate` | Transaction date/time |
| `UnitPrice` | Price per unit |
| `CustomerID` | Customer identifier |
| `Country` | Customer's country |
| `Revenue` | `Quantity × UnitPrice` |
| `Month` / `Year` | Derived from `InvoiceDate` |
| `OrderStatus` | Completed / Cancelled |
| `CustomerSegment` | New / Returning / Wholesale / Unknown |

---

## 📊 Pivot-Style Summary Tables

`Pivot_Summary` uses `SUMIFS` / `COUNTIFS` formulas (not native PivotTables) so every number
recalculates automatically whenever `Clean_Data` changes — Revenue by Country, Monthly Revenue,
and Revenue by Customer Segment, each with Orders, Units and Average Order Value.



---

## 🧮 Advanced Formulas

`Advanced_Formulas` demonstrates six core Excel formula skills side by side with their live
results, plus an interactive mini-panel (Selected Country / Selected Year / Filtered Revenue)
and a Country × Revenue ranking table using `RANK`, `LARGE` and `AVERAGE`.

| Formula Skill | Example Use |
|---|---|
| `XLOOKUP` / `VLOOKUP` | Look up a value by key, with an `IFERROR` fallback |
| `INDEX-MATCH` | Cross-sheet lookup into `Pivot_Summary` |
| `IF` | Simple threshold classification |
| Nested `IF` | Multi-tier classification (Bronze → Platinum) |
| `SUMIFS` | Conditional revenue totals by country + status |
| `COUNTIFS` | Conditional order counts by status |
| `RANK` / `LARGE` / `AVERAGE` | Country performance ranking (Top 3 / Above / Below Average) |



---

## 📈 Interactive Dashboard

The `Dashboard` sheet is the front page of the workbook: **KPI cards** (Total Revenue, Total
Orders, Average Order Value), a **Country revenue chart**, a **Segment breakdown chart**, and a
**monthly trend chart** — all driven off `Clean_Data`.

> Use the **Country** dropdown to filter the KPI and chart calculations. The project uses
> dropdown data-validation cells as a lightweight, portable alternative to native Excel slicers.


---

## 🗂️ Extra Pivot Tables

`pivot_table` provides three quick country/month breakdowns (order counts and revenue) as a
compact cross-check against the `Pivot_Summary` sheet.


## OUT SCREENSHORT
<img src="img/Screenshot 1.png">
<img src="img/Screenshot 2.png" >
<img src="img/Screenshot 3.png" >
<img src="img/Screenshot 4.png" >
<img src="" >

---

## 🛠️ Tech Stack

| Tool | Purpose |
|---|---|
| 🟢 **Microsoft Excel 2016+** | Core spreadsheet application |
| 🔎 **XLOOKUP / VLOOKUP / INDEX-MATCH** | Lookups across sheets |
| ➗ **SUMIFS / COUNTIFS** | Conditional aggregation |
| 🏆 **RANK / LARGE / AVERAGE** | Performance ranking |
| 📊 **Native Excel Charts** | Bar & column charts on the dashboard |
| ✅ **Data Validation (Dropdowns)** | Interactive Country/Year selectors |

---

## 🏆 Marks Mapping

| Component | Marks |
|---|---|
| Dataset | 2.5 |
| Cleaning | 5 |
| Pivot / Summary | 7.5 |
| Advanced Formulas | 7.5 |
| Visualization | 10 |
| Interactive Dashboard | 12.5 |
| Documentation | 5 |
| **Total** | **50** |

**Dataset source:** UCI Machine Learning Repository — *Online Retail* (UK online-retail
transactions: invoice, stock code, description, quantity, date, unit price, customer, country).

---

## 📄 License

This project is licensed under the **MIT License** — free to use, modify, and distribute with
attribution.

---

## 👤 Author

<div align="center">

### divyesh jadav


[![GitHub](https://img.shields.io/badge/GitHub-yourhandle-181717?style=for-the-badge&logo=github&logoColor=white)](https://github.com/isamaliya16)
[![LinkedIn](https://img.shields.io/badge/LinkedIn-Connect-0077B5?style=for-the-badge&logo=linkedin&logoColor=white)](https://www.linkedin.com/in/ayush-isamaliya-686533312/)

**🎓 Role:** Junior Data/Excel Analyst · Programming Enthusiast \
**📍 Location:** India \
**🛠️ Skills:** Excel · Data Cleaning · Pivot Analysis · Formulas · Dashboards

</div>

---

## 🙏 Acknowledgements

- 📚 [UCI Machine Learning Repository](https://archive.ics.uci.edu/) — Online Retail dataset
- 📊 [Microsoft Excel Support](https://support.microsoft.com/excel) — Function reference
- 🧮 [ExcelJet](https://exceljet.net/) — Formula patterns and examples

---

<div align="center">

---

*Last updated: 18 September, 2026*

</div>
