
# Sales Performance Dashboard - Power BI Project

## Structure
SalesPerformance.pbip - Main project file (open this in Power BI Desktop)
SalesPerformance.Dataset/ - Semantic model (TMDL format)
  - definition/model.tmdl - Model definition
  - definition/tables/*.tmdl - 5 tables: Sales, Products, Regions, Customers, Calendar
  - definition/relationships.tmdl - Relationships
SalesPerformance.Report/ - Report definition (PBIR format)
  - 3 pages: Sales Overview, Product Performance, Regional Analysis
Sales_Performance_Dataset.xlsx - Source Excel file (must be in same folder when refreshing)
PowerBI_DAX_Measures.txt - All DAX measures
Power_Query_M_Script.txt - Power Query M scripts

## How to Use
1. Ensure Power BI Desktop is updated to Jan 2024 or later (supports PBIP/TMDL)
2. Place Sales_Performance_Dataset.xlsx in same folder as .pbip file (or update M query path in TMDL files)
3. Double-click SalesPerformance.pbip - Power BI Desktop will open both dataset and report
4. If prompted for data source, point to Sales_Performance_Dataset.xlsx
5. Click Refresh to load 800 rows of transaction data
6. Measures are already defined in Sales table TMDL:
   - Total Sales, Total Profit, Total Orders, Avg Order Value, Profit Margin %, YoY Growth etc.
7. Build visuals:
   - Sales Overview: KPI cards (Total Sales, Profit, Orders, Margin), Line chart Monthly Trend, Bar Category, Pie Region
   - Product Performance: Top 10 products bar, Category vs Profit scatter
   - Regional Analysis: Map (Region), Channel performance, SalesRep table

## If PBIP doesn't open (older Power BI version):
1. Open Power BI Desktop -> Get Data -> Excel -> Select Sales_Performance_Dataset.xlsx
2. Import all 5 sheets
3. Create relationships as defined in relationships.tmdl
4. Copy DAX measures from PowerBI_DAX_Measures.txt
5. Recreate visuals as per Excel Dashboard sheet reference

## KPIs Implemented (no invented metrics)
All KPIs are derived from transaction-level Sales table using SUM, DISTINCTCOUNT, DIVIDE, CALCULATE - no hardcoded values.

## Files Included
- Excel Dashboard: ../Sales_Performance_Dashboard_Excel.xlsx
- This Power BI folder is self-contained and Git-friendly (PBIP format)

Generated: 2024 dataset (800 transactions, 12 products, 4 regions, 50 customers)
