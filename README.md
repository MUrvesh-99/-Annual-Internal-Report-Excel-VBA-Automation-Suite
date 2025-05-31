# Annual Internal Report Excel (VBA Macros)

## Project Overview
This project presents a fully automated VBA-based Excel reporting tool designed to generate an annual internal report covering sales, payments, and SLA compliance metrics. The solution was developed to streamline the analysis of performance trends across multiple regional offices using raw CSV inputs.

### Key Features
#### Automation Modules (Macros)
* Step 1: Load Data – Imports raw data from multiple CSVs (e.g., orders.csv, payments.csv).
* Step 2: Data Preparation – Cleans and organizes data into master tabs.
* Step 3A–3C: SLA Compliance – Populates, analyzes, and visualizes delivery performance vs. SLAs.
* Step 4: Payments Analysis – Summarizes monthly payment activity and values.
* Step 5: Sales Analysis – Analyzes office-wise revenue, profitability, and volume.
* Step 6: Summary Dashboard – Presents visual KPIs and insights for stakeholders.

#### Visual Insights
* Interactive dashboards
* SLA compliance heatmaps and charts
* Conditional formatting to highlight performance highs/lows

#### Required Input Files
Place the following CSVs in the same folder as the .xlsm file:

* orders.csv
* customers.csv
* employees.csv
* offices.csv
* orderdetails.csv
* payments.csv
* productlines.csv
* products.csv

### How to Use

* Open annual_internal_report.xlsm in Excel and enable macros.
* Go to the Refresh Report tab.
* Sequentially click each button: Step 1 → Step 2 → Step 3A ... Step 6.


### Notes
* Sheet names and workbook name must match exactly (case-sensitive).
* Macros are interdependent—do not run intermediate macros (e.g., Macro_99) directly.
* The Step 3C: SLA Visuals may be skipped if errors occur; it is not essential to final output.
