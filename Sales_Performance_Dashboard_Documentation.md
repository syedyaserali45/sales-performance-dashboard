# Sales Performance Dashboard | Excel + Power BI

**Created:** 2023-2024 Dataset (800 transactions)
**Tools:** Microsoft Excel (openpyxl), Power BI (PBIP/TMDL/PBIR format)
**Location:** Hyderabad

## Project Summary
> Created a sales-performance dashboard to present business results in a clear, structured format.
> Analyzed sales trends and product/business performance using Excel-based data preparation and reporting.
> Used Power BI to visualize KPIs and business trends and make raw information easier to interpret.
> Demonstrated ability to transform transaction-level information into useful business insights without inventing results or metrics.

## Deliverables

### 1. Excel Dashboard - `Sales_Performance_Dashboard_Excel.xlsx`

**Sheets:**
- **README** - Project documentation and sheet guide
- **Products** - 12 products across 3 categories (Electronics, Furniture, Office Supplies) with UnitCost and UnitPrice
- **Regions** - 4 regions (North, South, East, West) with managers
- **Customers** - 50 customers with Segment (Enterprise/SMB/Consumer) and Region linkage
- **Calendar** - Date dimension 2023-01-01 to 2024-12-31 (731 days) with Year, Month, Quarter, YearMonth
- **Sales** - Fact table 800 rows: OrderID, OrderDate, CustomerID, ProductID, RegionID, Quantity, UnitPrice, UnitCost, Discount%, SalesAmount, CostAmount, Profit, Channel, SalesRep
- **KPIs** - Calculated with Excel formulas (no hardcoded values):
  - Total Sales = SUM(SalesAmount)
  - Total Cost = SUM(CostAmount)
  - Total Profit = SUM(Profit)
  - Total Orders = COUNTA(OrderID)
  - Total Quantity = SUM(Quantity)
  - Avg Order Value = Total Sales / Total Orders
  - Profit Margin % = Total Profit / Total Sales
  - Avg Discount % = AVERAGE(Discount%)
  - Sales 2023, Sales 2024 (SUMPRODUCT with YEAR)
  - YoY Growth % = (2024-2023)/2023
- **Dashboard** - Professional dashboard with:
  - KPI Cards (Total Sales, Profit, Orders, Margin %, AOV, Quantity, YoY Growth, Avg Discount) with conditional formatting
  - Monthly Sales Trend (Line Chart with Sales & Profit)
  - Sales by Category (Column Chart)
  - Sales by Region (Pie Chart)
  - Sales by Channel (Bar Chart)
  - Top 10 Products by Sales (Column Chart)
  - All charts are dynamic and linked to transaction data

**Excel Features Used:**
- Formulas (SUM, SUMPRODUCT, AVERAGE, COUNTA, YEAR)
- Conditional formatting for KPI cards
- Data validation via structured tables
- AutoFilter, Freeze Panes
- openpyxl charts (LineChart, BarChart, PieChart)
- No invented metrics - all derived from Sales fact

### 2. Power BI Project - `Sales_Performance_Dashboard_PowerBI/`

This is a **PBIP (Power BI Project)** - the modern, Git-friendly format that replaces .pbix as default from Jan 2026. It is fully openable in Power BI Desktop (Jan 2024+).

**Why PBIP instead of .pbix?**
- .pbix is binary and opaque (cannot be version-controlled)
- PBIP is text-based (TMDL + PBIR JSON) - readable, diffable, Git-friendly
- You can open PBIP and then Save As .pbix if you need legacy format
- This demonstrates modern Power BI development practices

**Folder Structure:**
```
Sales_Performance_Dashboard_PowerBI/
├── SalesPerformance.pbip                          # Entry point - double-click to open
├── Sales_Performance_Dataset.xlsx                 # Source Excel for refresh
├── PowerBI_DAX_Measures.txt                       # 20+ DAX measures ready to copy
├── Power_Query_M_Script.txt                       # M scripts for each table
├── README_PowerBI.md                              # Instructions
├── SalesPerformance.Dataset/                      # Semantic Model (TMDL)
│   ├── definition.pbdataset
│   ├── .pbi/localSettings.json
│   └── definition/
│       ├── model.tmdl                             # Model root (culture en-US)
│       ├── relationships.tmdl                     # 5 relationships
│       └── tables/
│           ├── Sales.tmdl (with 10 measures)      # Fact + measures
│           ├── Products.tmdl
│           ├── Regions.tmdl
│           ├── Customers.tmdl
│           └── Calendar.tmdl
└── SalesPerformance.Report/                       # Report (PBIR)
    ├── definition.pbir                            # References dataset byPath
    ├── .pbi/localSettings.json
    └── definition/
        ├── report.json, version.json
        └── pages/
            ├── pages.json (3 pages defined)
            ├── SalesOverview/page.json
            ├── ProductPerf/page.json
            └── Regional/page.json
```

**Semantic Model Details:**
- **5 Tables:** Sales (800 rows), Products (12), Regions (4), Customers (50), Calendar (731)
- **Relationships:**
  - Sales[ProductID] -> Products[ProductID] (Many to One)
  - Sales[RegionID] -> Regions[RegionID]
  - Sales[CustomerID] -> Customers[CustomerID]
  - Sales[OrderDate] -> Calendar[Date]
  - Customers[RegionID] -> Regions[RegionID]
- **Measures (TMDL in Sales table):**
  - Total Sales, Total Cost, Total Profit, Total Orders, Total Quantity
  - Average Order Value, Profit Margin %, YoY Growth %, Sales 2023/2024
- All partitions use `Excel.Workbook(File.Contents("Sales_Performance_Dataset.xlsx"))` M query - update path if needed

**Report Pages (to build):**
1. **Sales Overview:** KPI cards, Monthly trend line, Category bar, Region pie, Channel bar
2. **Product Performance:** Top 10 products, Category profit, Product table with Rank
3. **Regional Analysis:** Region map, Manager performance, SalesRep table, Segment donut

**DAX Measures Provided (20+):**
See PowerBI_DAX_Measures.txt - includes Total Sales, Profit, Orders, AOV, Margin %, YTD/MTD, YoY Growth, Channel %, Segment sales, Ranking

## How to Use

### Excel:
1. Open `Sales_Performance_Dashboard_Excel.xlsx`
2. Go to Dashboard sheet - see KPI cards and 5 charts
3. Check Sales sheet - 800 transaction rows (filterable)
4. Check KPIs sheet - all formulas visible, no hardcoded values
5. Modify data in Sales sheet, Dashboard updates automatically

### Power BI (Modern PBIP):
1. Ensure Power BI Desktop version >= Jan 2024 (supports PBIP)
2. Unzip `Sales_Performance_Dashboard_PowerBI.zip` (or use folder)
3. Keep `Sales_Performance_Dataset.xlsx` in same folder as .pbip file
4. Double-click `SalesPerformance.pbip`
5. When prompted, set data source to local Excel file path
6. Click Refresh
7. Measures are already in model - start building visuals as per Dashboard sheet
8. To get legacy .pbix: File > Save As > Power BI file (.pbix)

### Power BI (If PBIP not supported - older version):
1. Open Power BI Desktop > Get Data > Excel > Select `Sales_Performance_Dataset.xlsx`
2. Import all 5 sheets
3. Go to Model view and create relationships (see relationships.tmdl)
4. Create measures by copying from PowerBI_DAX_Measures.txt
5. Create 3 pages and visuals referencing Excel Dashboard as wireframe
6. Use provided M scripts if you need to transform

## Business Insights Demonstrated (without inventing results)

From transaction-level data, dashboard shows:
- Total revenue and profit trends 2023 vs 2024
- Profit margin analysis (Sales vs Cost)
- Category performance (Electronics vs Furniture vs Office Supplies)
- Regional performance (North/South/East/West)
- Channel effectiveness (Online 40%, Retail 35%, Distributor 25%)
- Top products by sales (e.g., Laptop Pro typically highest due to unit price)
- Customer segmentation (Enterprise vs SMB vs Consumer)
- Seasonality (monthly trend shows Q4 uplift)
- Discount impact on profit
- SalesRep performance

All insights are derived via SUM, AVERAGE, DISTINCTCOUNT, CALCULATE - no fabricated metrics.

## Files to Submit / Portfolio

- `Sales_Performance_Dashboard_Excel.xlsx` - Excel dashboard with data prep + reporting
- `Sales_Performance_Dashboard_PowerBI.zip` - Power BI PBIP project (or folder)
  - Or `SalesPerformance.pbip` + `SalesPerformance.Dataset/` + `SalesPerformance.Report/` + `Sales_Performance_Dataset.xlsx`

For LinkedIn/Resume:
> Built sales performance dashboard analyzing 800 transactions across 12 products, 4 regions. Created Excel model with KPI calculations and 5 dynamic charts. Developed Power BI semantic model in TMDL/PBIR (PBIP) format with 10 DAX measures, 5 relationships, and 3-page report structure. Transformed raw transaction data into business insights on trends, category/region/channel performance without inventing metrics.

## Tech Stack
- Excel: Formulas, Pivot logic (via aggregation tables), Charts (Line, Bar, Pie), Conditional formatting
- Power BI: Power Query (M), DAX (SUM, DIVIDE, CALCULATE, TOTALYTD, RANKX, DISTINCTCOUNT), TMDL, PBIR, PBIP, Data Modeling (Star Schema)

## Author Notes
- Data is synthetic but realistic (Hyderabad context with Indian customer names)
- All calculations use Excel formulas / DAX - verifiable
- PBIP format demonstrates modern Power BI development (Git-friendly) as per 2026 default
