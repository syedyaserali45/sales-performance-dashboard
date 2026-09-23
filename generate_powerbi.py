import os
import json
import shutil
import pathlib

base = pathlib.Path("/home/user/Sales_Performance_Dashboard_PowerBI")
# Clean
if base.exists():
    shutil.rmtree(base)
base.mkdir(parents=True)

# 1. Create dataset Excel copy for Power BI source
# Copy from existing Excel file but create simplified version
import openpyxl
src_wb = openpyxl.load_workbook("/home/user/Sales_Performance_Dashboard_Excel.xlsx")
# Create new workbook for PBI dataset
wb = openpyxl.Workbook()
wb.remove(wb.active)
for sheet_name in ["Products", "Regions", "Customers", "Calendar", "Sales"]:
    src_ws = src_wb[sheet_name]
    dst_ws = wb.create_sheet(sheet_name)
    for row in src_ws.iter_rows(values_only=True):
        dst_ws.append(row)

dataset_excel_path = base / "Sales_Performance_Dataset.xlsx"
wb.save(dataset_excel_path)
print(f"Dataset Excel created: {dataset_excel_path}")

# 2. Create DAX measures file
dax_content = """
// ============================================
// Sales Performance Dashboard - DAX Measures
// ============================================
// Create these measures in your Sales table or dedicated _Measures table

Total Sales = SUM(Sales[SalesAmount])

Total Cost = SUM(Sales[CostAmount])

Total Profit = SUM(Sales[Profit])

Total Orders = DISTINCTCOUNT(Sales[OrderID])

Total Quantity = SUM(Sales[Quantity])

Average Order Value = DIVIDE([Total Sales], [Total Orders])

Profit Margin % = DIVIDE([Total Profit], [Total Sales])

Average Discount % = AVERAGE(Sales[Discount%])

// Time Intelligence
Sales 2023 = CALCULATE([Total Sales], FILTER(Calendar, Calendar[Year]=2023))

Sales 2024 = CALCULATE([Total Sales], FILTER(Calendar, Calendar[Year]=2024))

YoY Growth % = DIVIDE([Sales 2024] - [Sales 2023], [Sales 2023])

Sales YTD = TOTALYTD([Total Sales], Calendar[Date])

Sales MTD = TOTALMTD([Total Sales], Calendar[Date])

Profit YTD = TOTALYTD([Total Profit], Calendar[Date])

// Category & Region
Sales by Category = CALCULATE([Total Sales], ALLEXCEPT(Products, Products[Category]))

Top Product Sales = MAXX(TOPN(1, ADDCOLUMNS(VALUES(Products[ProductName]), "@Sales", [Total Sales]), [@Sales], DESC), [@Sales])

// Channel Performance
Online Sales % = DIVIDE(CALCULATE([Total Sales], Sales[Channel]="Online"), [Total Sales])

// Customer Segmentation
Enterprise Sales = CALCULATE([Total Sales], Customers[Segment]="Enterprise")

SMB Sales = CALCULATE([Total Sales], Customers[Segment]="SMB")

Consumer Sales = CALCULATE([Total Sales], Customers[Segment]="Consumer")

// Profitability
Profit per Order = DIVIDE([Total Profit], [Total Orders])

// Ranking
Product Rank = RANKX(ALL(Products[ProductName]), [Total Sales], , DESC, DENSE)

Region Rank = RANKX(ALL(Regions[RegionName]), [Total Sales], , DESC)
"""

with open(base / "PowerBI_DAX_Measures.txt", "w", encoding="utf-8") as f:
    f.write(dax_content)
print("DAX file created")

# 3. Power Query M scripts
m_content = """
// ============================================
// Power Query M Scripts for Sales Performance Dashboard
// ============================================

// --- Products ---
let
    Source = Excel.Workbook(File.Contents("Sales_Performance_Dataset.xlsx"), null, true),
    Products_Sheet = Source{[Item="Products",Kind="Sheet"]}[Data],
    PromotedHeaders = Table.PromoteHeaders(Products_Sheet, [PromoteAllScalars=true]),
    ChangedType = Table.TransformColumnTypes(PromotedHeaders,{
        {"ProductID", type text}, {"ProductName", type text}, {"Category", type text},
        {"UnitCost", Int64.Type}, {"UnitPrice", Int64.Type}
    })
in
    ChangedType

// --- Regions ---
let
    Source = Excel.Workbook(File.Contents("Sales_Performance_Dataset.xlsx"), null, true),
    Regions_Sheet = Source{[Item="Regions",Kind="Sheet"]}[Data],
    PromotedHeaders = Table.PromoteHeaders(Regions_Sheet, [PromoteAllScalars=true])
in
    PromotedHeaders

// --- Customers ---
let
    Source = Excel.Workbook(File.Contents("Sales_Performance_Dataset.xlsx"), null, true),
    Customers_Sheet = Source{[Item="Customers",Kind="Sheet"]}[Data],
    PromotedHeaders = Table.PromoteHeaders(Customers_Sheet, [PromoteAllScalars=true])
in
    PromotedHeaders

// --- Calendar ---
let
    Source = Excel.Workbook(File.Contents("Sales_Performance_Dataset.xlsx"), null, true),
    Calendar_Sheet = Source{[Item="Calendar",Kind="Sheet"]}[Data],
    PromotedHeaders = Table.PromoteHeaders(Calendar_Sheet, [PromoteAllScalars=true]),
    ChangedType = Table.TransformColumnTypes(PromotedHeaders,{
        {"Date", type date}, {"Year", Int64.Type}, {"Month", Int64.Type},
        {"MonthName", type text}, {"Quarter", type text}, {"YearMonth", type text}
    })
in
    ChangedType

// --- Sales (Fact) ---
let
    Source = Excel.Workbook(File.Contents("Sales_Performance_Dataset.xlsx"), null, true),
    Sales_Sheet = Source{[Item="Sales",Kind="Sheet"]}[Data],
    PromotedHeaders = Table.PromoteHeaders(Sales_Sheet, [PromoteAllScalars=true]),
    ChangedType = Table.TransformColumnTypes(PromotedHeaders,{
        {"OrderID", type text}, {"OrderDate", type date}, {"CustomerID", type text},
        {"ProductID", type text}, {"RegionID", type text}, {"Quantity", Int64.Type},
        {"UnitPrice", type number}, {"UnitCost", type number}, {"Discount%", type number},
        {"SalesAmount", type number}, {"CostAmount", type number}, {"Profit", type number},
        {"Channel", type text}, {"SalesRep", type text}
    }),
    // Add calculated columns
    AddedYear = Table.AddColumn(ChangedType, "Year", each Date.Year([OrderDate]), Int64.Type),
    AddedMonth = Table.AddColumn(AddedYear, "Month", each Date.Month([OrderDate]), Int64.Type)
in
    AddedMonth
"""

with open(base / "Power_Query_M_Script.txt", "w", encoding="utf-8") as f:
    f.write(m_content)
print("M script created")

# 4. Create PBIP structure
pbip_root = base
project_name = "SalesPerformance"
pbip_file = pbip_root / f"{project_name}.pbip"
dataset_folder = pbip_root / f"{project_name}.Dataset"
report_folder = pbip_root / f"{project_name}.Report"

# PBIP file
pbip_json = {
    "version": "1.0",
    "artifacts": [
        {"type": "dataset", "path": f"{project_name}.Dataset"},
        {"type": "report", "path": f"{project_name}.Report"}
    ],
    "settings": {}
}
with open(pbip_file, "w") as f:
    json.dump(pbip_json, f, indent=2)

# Dataset folders
(dataset_folder / ".pbi").mkdir(parents=True, exist_ok=True)
(dataset_folder / "definition" / "tables").mkdir(parents=True, exist_ok=True)

# definition.pbdataset
with open(dataset_folder / "definition.pbdataset", "w") as f:
    json.dump({
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/dataset/definitionProperties/1.0.0/schema.json",
        "version": "1.0"
    }, f, indent=2)

# localSettings
with open(dataset_folder / ".pbi" / "localSettings.json", "w") as f:
    json.dump({"version": "1.0"}, f, indent=2)

# model.tmdl
model_tmdl = """model SalesPerformance
    culture: en-US
    dataAccessOptions
        legacyRedirects
        returnErrorValuesAsNull
    defaultPowerBIDataSourceVersion: powerBI_V3
    sourceQueryCulture: en-US

    annotation PBI_ProTooling = "{\"isModelAuthoringEnabled\":false}"

    annotation __PBI_TimeIntelligenceEnabled = 0
"""

with open(dataset_folder / "definition" / "model.tmdl", "w", encoding="utf-8") as f:
    f.write(model_tmdl)

# relationships.tmdl
rel_tmdl = """relationship 11111111-1111-1111-1111-111111111111
    fromColumn: Sales.ProductID
    toColumn: Products.ProductID
    state: ready

relationship 22222222-2222-2222-2222-222222222222
    fromColumn: Sales.RegionID
    toColumn: Regions.RegionID
    state: ready

relationship 33333333-3333-3333-3333-333333333333
    fromColumn: Sales.CustomerID
    toColumn: Customers.CustomerID
    state: ready

relationship 44444444-4444-4444-4444-444444444444
    fromColumn: Sales.OrderDate
    toColumn: Calendar.Date
    state: ready

relationship 55555555-5555-5555-5555-555555555555
    fromColumn: Customers.RegionID
    toColumn: Regions.RegionID
    state: ready
"""

with open(dataset_folder / "definition" / "relationships.tmdl", "w", encoding="utf-8") as f:
    f.write(rel_tmdl)

# Helper to create table tmdl with M query referencing Excel
# Use relative path - user will need to update path
# We'll use parameter for file path

def make_table_tmdl(table_name, columns, measures=None):
    # columns: list of (name, datatype, sourceColumn)
    # measures: list of (name, expression)
    lines = [f"table {table_name}", f"    lineageTag: {table_name.lower()}-tag-0001", ""]
    # partition
    lines.append(f"    partition {table_name} = m")
    lines.append(f"        mode: import")
    lines.append(f"        source =")
    # M query - simplified
    m_query = f'''            let
                Source = Excel.Workbook(File.Contents("{dataset_excel_path.name}"), null, true),
                {table_name}_Sheet = Source{{[Item="{table_name}",Kind="Sheet"]}}[Data],
                #"Promoted Headers" = Table.PromoteHeaders({table_name}_Sheet, [PromoteAllScalars=true]),
                #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{{{{"Date", type date}}}})
            in
                #"Promoted Headers"'''
    # For simplicity use generic promoted headers
    if table_name == "Calendar":
        m_query = f'''            let
                Source = Excel.Workbook(File.Contents("{dataset_excel_path.name}"), null, true),
                {table_name}_Sheet = Source{{[Item="{table_name}",Kind="Sheet"]}}[Data],
                #"Promoted Headers" = Table.PromoteHeaders({table_name}_Sheet, [PromoteAllScalars=true])
            in
                #"Promoted Headers"'''
    else:
        m_query = f'''            let
                Source = Excel.Workbook(File.Contents("{dataset_excel_path.name}"), null, true),
                {table_name}_Sheet = Source{{[Item="{table_name}",Kind="Sheet"]}}[Data],
                #"Promoted Headers" = Table.PromoteHeaders({table_name}_Sheet, [PromoteAllScalars=true])
            in
                #"Promoted Headers"'''
    lines.append(m_query)
    lines.append("")
    for col_name, dtype, src in columns:
        lines.append(f"    column {col_name}")
        lines.append(f"        dataType: {dtype}")
        if src:
            lines.append(f"        sourceColumn: {src}")
        lines.append(f"        summarizeBy: none")
        lines.append("")
    if measures:
        for m_name, expr in measures:
            lines.append(f"    measure '{m_name}' = {expr}")
            lines.append(f"        formatString: #,##0.00")
            lines.append("")
    return "\n".join(lines)

# Products
products_cols = [
    ("ProductID", "string", "ProductID"),
    ("ProductName", "string", "ProductName"),
    ("Category", "string", "Category"),
    ("UnitCost", "int64", "UnitCost"),
    ("UnitPrice", "int64", "UnitPrice"),
]
with open(dataset_folder / "definition" / "tables" / "Products.tmdl", "w", encoding="utf-8") as f:
    f.write(make_table_tmdl("Products", products_cols))

# Regions
regions_cols = [
    ("RegionID", "string", "RegionID"),
    ("RegionName", "string", "RegionName"),
    ("Manager", "string", "Manager"),
    ("Zone", "string", "Zone"),
]
with open(dataset_folder / "definition" / "tables" / "Regions.tmdl", "w", encoding="utf-8") as f:
    f.write(make_table_tmdl("Regions", regions_cols))

# Customers
cust_cols = [
    ("CustomerID", "string", "CustomerID"),
    ("CustomerName", "string", "CustomerName"),
    ("Segment", "string", "Segment"),
    ("RegionID", "string", "RegionID"),
]
with open(dataset_folder / "definition" / "tables" / "Customers.tmdl", "w", encoding="utf-8") as f:
    f.write(make_table_tmdl("Customers", cust_cols))

# Calendar
cal_cols = [
    ("Date", "dateTime", "Date"),
    ("Year", "int64", "Year"),
    ("Month", "int64", "Month"),
    ("MonthName", "string", "MonthName"),
    ("Quarter", "string", "Quarter"),
    ("YearMonth", "string", "YearMonth"),
]
with open(dataset_folder / "definition" / "tables" / "Calendar.tmdl", "w", encoding="utf-8") as f:
    f.write(make_table_tmdl("Calendar", cal_cols))

# Sales with measures
sales_cols = [
    ("OrderID", "string", "OrderID"),
    ("OrderDate", "dateTime", "OrderDate"),
    ("CustomerID", "string", "CustomerID"),
    ("ProductID", "string", "ProductID"),
    ("RegionID", "string", "RegionID"),
    ("Quantity", "int64", "Quantity"),
    ("UnitPrice", "double", "UnitPrice"),
    ("UnitCost", "double", "UnitCost"),
    ("Discount%", "double", "Discount%"),
    ("SalesAmount", "double", "SalesAmount"),
    ("CostAmount", "double", "CostAmount"),
    ("Profit", "double", "Profit"),
    ("Channel", "string", "Channel"),
    ("SalesRep", "string", "SalesRep"),
]

sales_measures = [
    ("Total Sales", "SUM(Sales[SalesAmount])"),
    ("Total Cost", "SUM(Sales[CostAmount])"),
    ("Total Profit", "SUM(Sales[Profit])"),
    ("Total Orders", "DISTINCTCOUNT(Sales[OrderID])"),
    ("Total Quantity", "SUM(Sales[Quantity])"),
    ("Average Order Value", "DIVIDE([Total Sales], [Total Orders])"),
    ("Profit Margin %", "DIVIDE([Total Profit], [Total Sales])"),
    ("Sales 2023", "CALCULATE([Total Sales], FILTER(Calendar, Calendar[Year]=2023))"),
    ("Sales 2024", "CALCULATE([Total Sales], FILTER(Calendar, Calendar[Year]=2024))"),
    ("YoY Growth %", "DIVIDE([Sales 2024]-[Sales 2023],[Sales 2023])"),
]

with open(dataset_folder / "definition" / "tables" / "Sales.tmdl", "w", encoding="utf-8") as f:
    f.write(make_table_tmdl("Sales", sales_cols, sales_measures))

print("Dataset TMDL created")

# Report folder
(report_folder / ".pbi").mkdir(parents=True, exist_ok=True)
(report_folder / "definition" / "pages" / "SalesOverview").mkdir(parents=True, exist_ok=True)
(report_folder / "definition" / "pages" / "ProductPerf").mkdir(parents=True, exist_ok=True)
(report_folder / "definition" / "pages" / "Regional").mkdir(parents=True, exist_ok=True)
(report_folder / "StaticResources" / "RegisteredResources").mkdir(parents=True, exist_ok=True)

# definition.pbir
with open(report_folder / "definition.pbir", "w") as f:
    json.dump({
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definitionProperties/2.0.0/schema.json",
        "version": "4.0",
        "datasetReference": {
            "byPath": {"path": f"../{project_name}.Dataset"}
        }
    }, f, indent=2)

with open(report_folder / ".pbi" / "localSettings.json", "w") as f:
    json.dump({"version":"1.0"}, f, indent=2)

# definition/report.json
report_json = {
    "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/report/2.0.0/schema.json",
    "themeCollection": {
        "baseTheme": {
            "name": "CY24SU10",
            "reportVersionAtImport": "5.81",
            "type": 2
        }
    },
    "resourcePackages": []
}
with open(report_folder / "definition" / "report.json", "w") as f:
    json.dump(report_json, f, indent=2)

# version.json
with open(report_folder / "definition" / "version.json", "w") as f:
    json.dump({"version":"1.0"}, f, indent=2)

# pages/pages.json
pages_json = {
    "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/pages/2.0.0/schema.json",
    "activePageName": "SalesOverview",
    "pages": [
        {"displayName": "Sales Overview", "name": "SalesOverview", "displayOption": 1},
        {"displayName": "Product Performance", "name": "ProductPerf", "displayOption": 1},
        {"displayName": "Regional Analysis", "name": "Regional", "displayOption": 1}
    ]
}
with open(report_folder / "definition" / "pages" / "pages.json", "w") as f:
    json.dump(pages_json, f, indent=2)

# page.json for each page
for page_name, display_name in [("SalesOverview","Sales Overview"), ("ProductPerf","Product Performance"), ("Regional","Regional Analysis")]:
    page_json = {
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/page/2.0.0/schema.json",
        "displayName": display_name,
        "name": page_name,
        "displayOption": 1,
        "height": 720,
        "width": 1280,
        "objects": {}
    }
    with open(report_folder / "definition" / "pages" / page_name / "page.json", "w") as f:
        json.dump(page_json, f, indent=2)

print("Report structure created")

# Create README for Power BI folder
readme_pb = f"""
# Sales Performance Dashboard - Power BI Project

## Structure
{project_name}.pbip - Main project file (open this in Power BI Desktop)
{project_name}.Dataset/ - Semantic model (TMDL format)
  - definition/model.tmdl - Model definition
  - definition/tables/*.tmdl - 5 tables: Sales, Products, Regions, Customers, Calendar
  - definition/relationships.tmdl - Relationships
{project_name}.Report/ - Report definition (PBIR format)
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
"""

with open(base / "README_PowerBI.md", "w", encoding="utf-8") as f:
    f.write(readme_pb)

print("Power BI project created at:", base)
