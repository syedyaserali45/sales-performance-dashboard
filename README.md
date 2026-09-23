# Sales Performance Dashboard

A sales analytics project built in Excel and Power BI to analyze transaction-level sales data, track KPIs, and present business performance across regions, products, channels, and customers.

## Overview

This project turns raw sales data into a clear reporting package for business review. It includes:

- An Excel dashboard with KPI cards and charts
- A Power BI project in modern PBIP format
- Source data and documentation for the dashboard model
- DAX and Power Query assets for reusable analytics

## Project contents

- [Sales_Performance_Dashboard_Excel.xlsx](Sales_Performance_Dashboard_Excel.xlsx) — Excel dashboard and supporting data sheets
- [Sales_Performance_Dashboard_PowerBI.zip](Sales_Performance_Dashboard_PowerBI.zip) — packaged Power BI project assets
- [Sales_Performance_Dashboard_PowerBI/](Sales_Performance_Dashboard_PowerBI/) — PBIP project folder with dataset, model, and report definitions
- [Sales_Performance_Dashboard_Documentation.md](Sales_Performance_Dashboard_Documentation.md) — detailed project documentation
- [generate_excel.py](generate_excel.py) — Python script to generate the Excel workbook data/model
- [generate_powerbi.py](generate_powerbi.py) — Python script to generate Power BI-related assets

## Key business questions addressed

- What are total sales, profit, margin, and order trends?
- Which products, regions, and channels perform best?
- How does year-over-year performance look across 2023 and 2024?
- Which customer segments drive the most value?
- How can sales data be transformed into a reusable dashboard and reporting model?

## Tools used

- Microsoft Excel
- Power BI (PBIP / TMDL / PBIR)
- Python for data generation and automation
- Power Query and DAX for dashboard logic

## Folder structure

```text
Sales/
├── README.md
├── .gitignore
├── generate_excel.py
├── generate_powerbi.py
├── Sales_Performance_Dashboard_Excel.xlsx
├── Sales_Performance_Dashboard_PowerBI/
├── Sales_Performance_Dashboard_PowerBI.zip
├── Sales_Performance_Dashboard_Documentation.md
└── Unselected files/
```

## How to use

### Excel
1. Open [Sales_Performance_Dashboard_Excel.xlsx](Sales_Performance_Dashboard_Excel.xlsx)
2. Review the Dashboard sheet for KPI summaries and charts
3. Use the underlying data sheets to inspect transaction-level detail

### Power BI
1. Extract or open the project in [Sales_Performance_Dashboard_PowerBI/](Sales_Performance_Dashboard_PowerBI/)
2. Open the `.pbip` project in Power BI Desktop
3. Refresh the data source if needed

## Notes

This project is designed as a portfolio-style business intelligence dashboard using realistic transactional data and modern BI tooling. It is suitable for showcasing Excel and Power BI reporting capabilities in a clean GitHub repository.

## License

This project is provided for educational and portfolio use. If a license file is added later, it should be referenced here.
