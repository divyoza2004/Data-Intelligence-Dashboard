# Data Intelligence Dashboard – Final Project

A full analysis and reporting workbook built in Excel (generated with Python / openpyxl) on a dataset of 200+ customer transactions.

**Deliverable:** `Data_Intelligence_Dashboard_Solutions.xlsx`

---

## 1. Objective

Create a full-fledged analysis and reporting solution using Excel to analyze customer transactions, find insights, and build visualizations and dashboards.

The workbook must:

- Be well structured, with multiple sheets (Raw Data, Analysis, Visualizations, Dashboard).
- Give summary insights with key metrics: **Total Revenue, Top Customers, Most Sold Products, Regional Sales**.
- Use every required formula and function listed in the project brief.

---

## ▶ Demo Video

<a href="https://drive.google.com/file/d/1SsEeLGzhFWsQl6_RjEt73qVriQXatfwM/view?usp=sharing" target="_blank" rel="noopener noreferrer">
  <img src="https://img.shields.io/badge/▶-Watch%20Demo%20Video-181717?style=for-the-badge&logo=github&logoColor=white" alt="Watch Demo Video" />
</a>


## 2. Dataset

Sheet **Raw Data** – 220 transaction records, 14 columns.

| Column | Field | Description |
|---|---|---|
| A | Transaction_ID | Unique ID (TRX0001 – TRX0250) |
| B | Date | Purchase date |
| C | Customer_ID | CUST001 – CUST050 |
| D | Customer_Name | Customer full name |
| E | Product_ID | P001 – P010 |
| F | Product_Name | Laptop, Smartphone, Monitor, etc. |
| G | Category | Electronics / Appliances / Furniture |
| H | Quantity | Units bought (1–5) |
| I | Unit_Price | Price per unit |
| J | Payment_Method | Cash / Credit Card / Debit Card / PayPal |
| K | Region | North / South / East / West / Central |
| L | Customer_Segment | Basic / Standard / Premium |
| M | Customer_Since | Customer registration date |
| N | Total_Amount | `=Quantity * Unit_Price` (formula) |

> The first 26 rows come from the project screenshot. The remaining rows are generated sample data. Paste your real data over it (keep columns A–N) and every sheet recalculates.

---

## 3. Features

- 13 solution sheets, one per required topic (Q1–Q13).
- Everything is formula-driven, so results update when Raw Data changes.
- Conditional formatting with icon sets, data bars and colour rules.
- Four dashboard charts (bar, line, pie, region bar) and one regression chart with a trendline.
- KPI cards and KPI arrow indicators.
- Auto-written insight sentences (text formulas).
- Scenario analysis, Goal Seek set-up and a self-freezing timestamp.
- Works in all Excel versions (a 365 `FILTER` alternative is also given).

---

## 4. Requirements

### Software

| Requirement | Purpose |
|---|---|
| Microsoft Excel 2016+ (365 recommended) | Open and use the workbook |
| Python 3.9+ | Run the generator script |
| `openpyxl` | Build the workbook (`pip install openpyxl`) |
| LibreOffice (optional) | Recalculate formulas / check for errors |

### Excel settings

| Setting | Needed for |
|---|---|
| File > Options > Formulas > **Enable iterative calculation** | Q4 timestamp (already switched on in the file) |
| File > Options > Add-ins > **Analysis ToolPak** | Q6 regression (optional) |

### Run the script

```bash
pip install openpyxl
python build.py
```

Output: `Data_Intelligence_Dashboard_Solutions.xlsx`

---

## 5. Workbook structure

| Sheet | Role |
|---|---|
| Raw Data | Source transactions |
| Q1_DateTime … Q8_HighValue | Analysis |
| Q9_Dashboard | Visualizations and dashboard |
| Q10_MostFrequent … Q12_TextFunctions | Analysis |
| Q13_FinalReport | Final report and insights |

---

## 6. Tasks explained

### Task summary table

| # | Task | Excel functions / tools | Sheet | Result |
|---|---|---|---|---|
| 1 | Date & Time functions | `TODAY`, `NOW`, `DATEDIF`, `EOMONTH` | Q1_DateTime | Days/months since purchase, customer tenure, month-end dates |
| 2 | FILTER for multi-value returns | `FILTER` (365) or `INDEX` + `SMALL` + `IF` | Q2_Filter | All rows matching Category, Region and minimum amount |
| 3 | Conditional formatting with icons/arrows | Icon sets, data bars, formula rules | Q3_CondFormat | Arrows and traffic lights on amount and quantity |
| 4 | Timestamp using NOW and iteration | `=IF(A8<>"",IF(B8="",NOW(),B8),"")` | Q4_Timestamp | Time is captured once and then frozen |
| 5 | What-If analysis | Scenario Manager, Goal Seek | Q5_WhatIf | Worst/Base/Best profit and price change needed for a target |
| 6 | Linear regression | `SLOPE`, `INTERCEPT`, `RSQ`, `STEYX`, `FORECAST`, `LINEST` | Q6_Regression | Monthly revenue trend, forecast and trendline chart |
| 7 | Data analysis and storytelling | `SUMIF`, `COUNTIF`, `TEXT`, `INDEX/MATCH` | Q7_Storytelling | Metrics, cuts by category/payment/segment/region, insights |
| 8 | GroupBy to find high-value customers | `SUMIF`, `COUNTIF`, `RANK`, `PERCENTILE`, `LARGE` | Q8_HighValue | Spend per customer, top 10, High-Value tag |
| 9 | Dashboard | Summary tables, charts, KPI icons | Q9_Dashboard | KPI cards, bar/line/pie charts |
| 10 | Most frequent product / repeating names | `COUNTIF`, `MODE`, `MATCH`, `INDEX` | Q10_MostFrequent | Most purchased product, repeat customers |
| 11 | Compare two lists, extract matches | `COUNTIF`, `IF`, `VLOOKUP` | Q11_CompareLists | Names present in both lists |
| 12 | Abbreviations with TEXT functions | `LEFT`, `MID`, `FIND`, `RIGHT`, `UPPER`, `PROPER`, `LEN`, `TEXT` | Q12_TextFunctions | Initials, short names, product/region/invoice codes |
| 13 | Final report with insights | Links to previous sheets, `TEXT` | Q13_FinalReport | Key metrics, top customers, regional sales, recommendations |

### Task details

**Q1 – Date & Time Functions.** `TODAY()` returns today's date and `NOW()` the current date and time. `DATEDIF(start, end, "D"/"M"/"Y"/"YM")` gives the gap between dates, used here for days since purchase and customer tenure. `EOMONTH(date, 0)` returns the month-end date and `EOMONTH(date, 1)` the next month-end.

**Q2 – FILTER.** Change the yellow cells (Category, Region, Minimum Amount) and the list of matching transactions updates. A helper column finds the row numbers of matches and `INDEX` pulls each field. The one-cell 365 formula is written at the bottom of the sheet.

**Q3 – Conditional formatting.** Total_Amount gets a three-arrow icon set, Quantity gets traffic lights, and a formula-based arrow (green ▲ above average, red ▼ below) compares each amount with the average.

**Q4 – Timestamp.** The formula refers to its own cell (a circular reference) so `NOW()` is captured once when you type in column A and then stays fixed. This needs iterative calculation turned on.

**Q5 – What-If analysis.** Part A compares Worst/Base/Best scenarios for price, volume and cost. Part B is a live model with changing cells B19:B21 for Scenario Manager. Part C is a Goal Seek task (set B24 to the target by changing B19) with the formula answer in B28 for checking.

**Q6 – Linear regression.** Monthly revenue (Y) is regressed on month number (X). The sheet reports slope, intercept, R², correlation, standard error and forecasts for the next three months, with a scatter chart and trendline. For the full ToolPak output, use Data > Data Analysis > Regression with Y = C6:C17 and X = B6:B17.

**Q7 – Data analysis and storytelling.** Key metrics (revenue, orders, units, average order value, customers) and breakdowns by category, payment method, segment and region. Five insight sentences are written by formulas and update with the data.

**Q8 – High-value customers.** Transactions are grouped per customer with `SUMIF` and `COUNTIF`, ranked, and tagged **High-Value** when spend reaches the 80th percentile. A top-10 table is included. In Microsoft 365 the same result can be had with `GROUPBY`.

**Q9 – Dashboard.** Five KPI cards, summary tables that work like pivot tables, month-over-month arrows, and four charts. To add real PivotTables, slicers and a timeline: select Raw Data > Insert > PivotTable, then PivotTable Analyze > Insert Slicer (Region, Category) and Insert Timeline (Date).

**Q10 – Most frequent.** Counts purchases per product and finds the most frequent one with `INDEX/MATCH` on `MAX`, and with `INDEX(range, MODE(MATCH(range, range, 0)))`. A customer table flags repeat buyers.

**Q11 – Compare two lists.** List A is customers who bought Electronics, List B is customers who bought Furniture. The matching column returns a name only when it exists in both lists.

**Q12 – TEXT functions.** Builds first/last names, initials, a short name (e.g. `PAU.BAK`), product code, region code, a date label, and an invoice code such as `INV-2404-0155-MONC`.

**Q13 – Final report.** Brings together total revenue, top customer, most sold product, top region, top 5 customers, product and regional tables, a written story, and recommendations.

---

## 7. Python generation approach

| Step | openpyxl feature | Used for |
|---|---|---|
| Create workbook | `Workbook()` | Start the file |
| Write formulas | `ws["B5"] = "=SUM(...)"` | All calculations stay live in Excel |
| Array formulas | `ArrayFormula(ref, text)` | Q2 helper, Q10 `MODE` |
| Conditional formatting | `IconSetRule`, `DataBarRule`, `FormulaRule` | Q3, Q8, Q9 |
| Charts | `BarChart`, `LineChart`, `PieChart`, `ScatterChart` | Q6, Q9, Q13 |
| Trendline | `Trendline(trendlineType="linear")` | Q6 |
| Data validation | `DataValidation(type="list")` | Q5 scenario picker |
| Iterative calc | `CalcProperties(iterate=True)` | Q4 timestamp |
| Styling | `Font`, `PatternFill`, `Border` | Green question banners, headers |


## 8. Key results (sample data)

| Metric | Value |
|---|---|
| Total Revenue | 194,873.62 |
| Transactions | 220 |
| Average Order Value | about 886 |
| Top Category | Electronics (about 72% of revenue) |
| Top Customer | Kathleen Bennett (10,229.72) |
| Top Region | North |
| High-Value Customers | 10 customers, about 41% of revenue |

These figures come from the sample dataset and will change with your real data.

---

## 9. Limitations

- True PivotTables, slicers, timelines, Scenario Manager and ToolPak output cannot be saved from Python. The workbook gives equivalent formula tables plus the steps to create them.
- The Q6 trend is weak (R² near 0) because the extra rows are random sample data.
- Q4 timestamps stamp correctly in Excel desktop; some previewers show 00:00.
