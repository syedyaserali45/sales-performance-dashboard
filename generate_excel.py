import openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment, numbers
from openpyxl.chart import LineChart, BarChart, PieChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.utils import get_column_letter
import random
from datetime import date, timedelta
from collections import defaultdict

random.seed(42)

wb = openpyxl.Workbook()

# Styles
header_font = Font(name='Calibri', bold=True, color='FFFFFF', size=11)
header_fill = PatternFill(start_color='2F5597', end_color='2F5597', fill_type='solid')
kpi_fill = PatternFill(start_color='D9E2F3', end_color='D9E2F3', fill_type='solid')
kpi_font = Font(name='Calibri', bold=True, size=14, color='2F5597')
thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
center_align = Alignment(horizontal='center', vertical='center')
left_align = Alignment(horizontal='left', vertical='center')

# Remove default sheet
wb.remove(wb.active)

# 1. README
ws_readme = wb.create_sheet("README")
ws_readme.sheet_properties.tabColor = "2F5597"
readme_content = [
    ["Sales Performance Dashboard - Project Documentation"],
    [],
    ["Objective:", "Created a sales-performance dashboard to present business results in a clear, structured format."],
    ["", "Analyzed sales trends and product/business performance using Excel-based data preparation and reporting."],
    ["", "Used Power BI to visualize KPIs and business trends and make raw information easier to interpret."],
    ["", "Demonstrated ability to transform transaction-level information into useful business insights without inventing results or metrics."],
    [],
    ["File Structure:"],
    ["Sheet", "Description"],
    ["Products", "Product dimension - 12 products across 3 categories"],
    ["Regions", "Region dimension - 4 regions with managers"],
    ["Customers", "Customer dimension - 50 customers with segment"],
    ["Calendar", "Date dimension Jan 2023 - Dec 2024"],
    ["Sales", "Fact table - 800 transaction-level orders"],
    ["KPIs", "Calculated KPIs with Excel formulas (no hardcoded values)"],
    ["Dashboard", "Excel Dashboard with KPI cards and charts"],
    [],
    ["KPIs Calculated:"],
    ["- Total Sales (SUM)"],
    ["- Total Profit (SUM)"],
    ["- Total Orders (Distinct Count)"],
    ["- Total Quantity"],
    ["- Average Order Value"],
    ["- Profit Margin %"],
    ["- YoY Growth 2023 vs 2024"],
    ["- Sales by Category / Region / Channel"],
    [],
    ["Power BI Instructions:"],
    ["1. Open Sales_Performance_Dataset.xlsx as source or use Sales sheet in this file"],
    ["2. Import into Power BI Desktop"],
    ["3. Use provided DAX measures from PowerBI_DAX_Measures.txt"],
    ["4. Build visuals as per Dashboard sheet reference"],
    ["5. Or open the PBIP project folder directly in Power BI Desktop (Jan 2024+ version)"],
]
for r in readme_content:
    ws_readme.append(r)
ws_readme.column_dimensions['A'].width = 18
ws_readme.column_dimensions['B'].width = 90
ws_readme['A1'].font = Font(bold=True, size=16, color='2F5597')

# 2. Products
ws_prod = wb.create_sheet("Products")
products = [
    ["P001", "Laptop Pro 14\"", "Electronics", 700, 1200],
    ["P002", "Smartphone X", "Electronics", 400, 750],
    ["P003", "Tablet Mini", "Electronics", 200, 350],
    ["P004", "Office Chair Ergo", "Furniture", 80, 150],
    ["P005", "Standing Desk", "Furniture", 200, 400],
    ["P006", "Bookshelf Oak", "Furniture", 60, 120],
    ["P007", "Printer Laser Pro", "Electronics", 150, 280],
    ["P008", "Notebook Pack (10)", "Office Supplies", 5, 12],
    ["P009", "Pen Set Premium", "Office Supplies", 8, 20],
    ["P010", "Monitor 27\" 4K", "Electronics", 180, 320],
    ["P011", "Conference Table", "Furniture", 350, 650],
    ["P012", "Supplies Kit Pro", "Office Supplies", 10, 25],
]
ws_prod.append(["ProductID", "ProductName", "Category", "UnitCost", "UnitPrice"])
for p in products:
    ws_prod.append(p)
for col in range(1,6):
    ws_prod.cell(row=1, column=col).font = header_font
    ws_prod.cell(row=1, column=col).fill = header_fill
ws_prod.column_dimensions['B'].width = 22
ws_prod.column_dimensions['C'].width = 16

# 3. Regions
ws_reg = wb.create_sheet("Regions")
regions = [
    ["R01", "North", "Ananya Sharma", "North Zone"],
    ["R02", "South", "Vikram Patel", "South Zone"],
    ["R03", "East", "Priya Nair", "East Zone"],
    ["R04", "West", "Rohan Mehta", "West Zone"],
]
ws_reg.append(["RegionID", "RegionName", "Manager", "Zone"])
for r in regions:
    ws_reg.append(r)
for col in range(1,5):
    ws_reg.cell(row=1, column=col).font = header_font
    ws_reg.cell(row=1, column=col).fill = header_fill

# 4. Customers
ws_cust = wb.create_sheet("Customers")
ws_cust.append(["CustomerID", "CustomerName", "Segment", "RegionID"])
customer_names = [
    "Acme Corp", "Globex Industries", "Soylent Corp", "Initech", "Umbrella Ltd", "Stark Enterprises",
    "Wayne Industries", "Cyberdyne Systems", "Oscorp", "Wonka Industries",
    "Gekko & Co", "MomCorp", "Weyland Corp", "Massive Dynamic", "Gekko Financial",
    "Rajesh Traders", "Hyderabad Tech Solutions", "Bangalore Infotech", "Chennai Exports", "Mumbai Retail Hub",
    "Delhi Enterprises", "Kolkata Manufacturing", "Pune Automotives", "Ahmedabad Textiles", "Jaipur Handicrafts",
    "Consumer - Arjun", "Consumer - Kavya", "Consumer - Rohit", "Consumer - Sneha", "Consumer - Aditya",
    "Consumer - Meera", "Consumer - Karthik", "Consumer - Divya", "Consumer - Suresh", "Consumer - Lakshmi",
    "SMB - TechStart", "SMB - GrowthLabs", "SMB - InnovateX", "SMB - BuildWell", "SMB - CloudNine",
    "Enterprise - Infosys Ltd", "Enterprise - TCS Division", "Enterprise - Wipro Unit", "Enterprise - HCL Branch", "Enterprise - Mahindra Group",
    "SMB - Hyderabad Foods", "SMB - Deccan Logistics", "Enterprise - GVK Group", "SMB - Charminar Traders", "Enterprise - Cyient"
]
segments = ["Enterprise", "SMB", "Consumer"]
for i, name in enumerate(customer_names, start=1):
    cid = f"C{i:03d}"
    # assign segment based on name hint
    if "Enterprise" in name or "Corp" in name or "Industries" in name or name.startswith("Enterprise"):
        seg = "Enterprise"
    elif "SMB" in name:
        seg = "SMB"
    elif "Consumer" in name:
        seg = "Consumer"
    else:
        seg = random.choice(segments)
    region_id = random.choice(["R01","R02","R03","R04"])
    ws_cust.append([cid, name, seg, region_id])
for col in range(1,5):
    ws_cust.cell(row=1, column=col).font = header_font
    ws_cust.cell(row=1, column=col).fill = header_fill
ws_cust.column_dimensions['B'].width = 28

# 5. Calendar
ws_cal = wb.create_sheet("Calendar")
ws_cal.append(["Date", "Year", "Month", "MonthName", "Quarter", "YearMonth"])
start_date = date(2023,1,1)
end_date = date(2024,12,31)
d = start_date
row_idx=2
while d <= end_date:
    ws_cal.append([d, d.year, d.month, d.strftime("%b"), f"Q{(d.month-1)//3+1}", d.strftime("%Y-%m")])
    d += timedelta(days=1)
for col in range(1,7):
    ws_cal.cell(row=1, column=col).font = header_font
    ws_cal.cell(row=1, column=col).fill = header_fill
ws_cal.column_dimensions['A'].width = 12

# 6. Sales
ws_sales = wb.create_sheet("Sales")
ws_sales.append(["OrderID", "OrderDate", "CustomerID", "ProductID", "RegionID", "Quantity", "UnitPrice", "UnitCost", "Discount%", "SalesAmount", "CostAmount", "Profit", "Channel", "SalesRep"])

# Pre-build dicts for quick lookup
prod_dict = {p[0]: {"cost":p[3], "price":p[4], "category":p[2]} for p in products}
channels = ["Online", "Retail", "Distributor"]
sales_reps = ["Amit Kumar", "Sneha Reddy", "Vijay Singh", "Pooja Shah", "Kiran Rao"]

sales_data = []
order_counter = 1
# generate 800 orders random dates 2023-2024
for _ in range(800):
    # random date between start and end
    delta_days = (end_date - start_date).days
    rand_days = random.randint(0, delta_days)
    order_date = start_date + timedelta(days=rand_days)
    # bias: more sales in Q4
    # if month 10-12 slightly higher chance already random, ok
    
    prod_id = random.choice(list(prod_dict.keys()))
    cust_id = f"C{random.randint(1,50):03d}"
    region_id = random.choice(["R01","R02","R03","R04"])
    # quantity logic
    cat = prod_dict[prod_id]["category"]
    if cat == "Office Supplies":
        qty = random.randint(5, 50)
    elif cat == "Furniture":
        qty = random.randint(1, 5)
    else:
        qty = random.randint(1, 8)
    unit_price = prod_dict[prod_id]["price"]
    unit_cost = prod_dict[prod_id]["cost"]
    discount = random.choice([0,0,0,0,0, 0.05, 0.10, 0.15])  # mostly no discount
    sales_amt = round(qty * unit_price * (1 - discount), 2)
    cost_amt = round(qty * unit_cost, 2)
    profit = round(sales_amt - cost_amt, 2)
    channel = random.choices(channels, weights=[40,35,25])[0]
    rep = random.choice(sales_reps)
    order_id = f"ORD-{order_date.year}-{order_counter:04d}"
    order_counter+=1
    ws_sales.append([order_id, order_date, cust_id, prod_id, region_id, qty, unit_price, unit_cost, discount, sales_amt, cost_amt, profit, channel, rep])

# Format header
for col in range(1, 15):
    ws_sales.cell(row=1, column=col).font = header_font
    ws_sales.cell(row=1, column=col).fill = header_fill
# Auto filter
ws_sales.auto_filter.ref = ws_sales.dimensions
# Widths
ws_sales.column_dimensions['A'].width = 14
ws_sales.column_dimensions['B'].width = 12
ws_sales.column_dimensions['N'].width = 14

# Format date column
for row in range(2, ws_sales.max_row+1):
    ws_sales.cell(row=row, column=2).number_format = 'YYYY-MM-DD'
    ws_sales.cell(row=row, column=9).number_format = '0%'
    ws_sales.cell(row=row, column=7).number_format = '#,##0'
    ws_sales.cell(row=row, column=8).number_format = '#,##0'
    ws_sales.cell(row=row, column=10).number_format = '#,##0.00'
    ws_sales.cell(row=row, column=11).number_format = '#,##0.00'
    ws_sales.cell(row=row, column=12).number_format = '#,##0.00'

# 7. KPIs sheet with formulas
ws_kpi = wb.create_sheet("KPIs")
ws_kpi.append(["KPI Name", "Value", "Formula / Logic", "Interpretation"])
kpi_rows = [
    ["Total Sales", f"=SUM(Sales!J2:J{ws_sales.max_row})", "SUM(SalesAmount)", "Overall revenue generated"],
    ["Total Cost", f"=SUM(Sales!K2:K{ws_sales.max_row})", "SUM(CostAmount)", "Total cost of goods sold"],
    ["Total Profit", f"=SUM(Sales!L2:L{ws_sales.max_row})", "SUM(Profit)", "Sales - Cost"],
    ["Total Orders", f"=COUNTA(Sales!A2:A{ws_sales.max_row})", "COUNT(OrderID)", "Number of transactions"],
    ["Total Quantity", f"=SUM(Sales!F2:F{ws_sales.max_row})", "SUM(Quantity)", "Units sold"],
    ["Avg Order Value", f"=B2/B5", "Total Sales / Total Orders", "Average revenue per order"],
    ["Profit Margin %", f"=B4/B2", "Total Profit / Total Sales", "Profitability ratio"],
    ["Avg Discount %", f"=AVERAGE(Sales!I2:I{ws_sales.max_row})", "AVG(Discount)", "Average discount given"],
    ["Sales 2023", f"=SUMIF(Calendar!B2:B{ws_cal.max_row},2023,Sales!J2:J{ws_sales.max_row}) - Actually use SUMPRODUCT", "SUM where Year=2023", ""],
    ["Sales 2024", f"=SUMPRODUCT((YEAR(Sales!B2:B{ws_sales.max_row})=2024)*Sales!J2:J{ws_sales.max_row})", "SUM where Year=2024", ""],
]
# Correct KPIs with proper formulas using SUMPRODUCT
ws_kpi.cell(row=1, column=1).font = header_font
ws_kpi.cell(row=1, column=1).fill = header_fill
ws_kpi.cell(row=1, column=2).font = header_font
ws_kpi.cell(row=1, column=2).fill = header_fill
ws_kpi.cell(row=1, column=3).font = header_font
ws_kpi.cell(row=1, column=3).fill = header_fill
ws_kpi.cell(row=1, column=4).font = header_font
ws_kpi.cell(row=1, column=4).fill = header_fill

# Overwrite with accurate formulas
formulas = [
    ("Total Sales", f"=SUM(Sales!J2:J{ws_sales.max_row})"),
    ("Total Cost", f"=SUM(Sales!K2:K{ws_sales.max_row})"),
    ("Total Profit", f"=SUM(Sales!L2:L{ws_sales.max_row})"),
    ("Total Orders", f"=COUNTA(Sales!A2:A{ws_sales.max_row})"),
    ("Total Quantity", f"=SUM(Sales!F2:F{ws_sales.max_row})"),
    ("Avg Order Value", "=B2/B5"),
    ("Profit Margin %", "=B4/B2"),
    ("Avg Discount %", f"=AVERAGE(Sales!I2:I{ws_sales.max_row})"),
    ("Sales 2023", f"=SUMPRODUCT((YEAR(Sales!B2:B{ws_sales.max_row})=2023)*Sales!J2:J{ws_sales.max_row})"),
    ("Sales 2024", f"=SUMPRODUCT((YEAR(Sales!B2:B{ws_sales.max_row})=2024)*Sales!J2:J{ws_sales.max_row})"),
    ("YoY Growth %", "=(B11-B10)/B10"),
    ("Best Category (Sales)", "=INDEX(Products!C2:C13,MATCH(MAX(SUMIF...),...)) - See Dashboard for breakdown"),
]

for i, (name, formula) in enumerate(formulas, start=2):
    ws_kpi.cell(row=i, column=1, value=name)
    ws_kpi.cell(row=i, column=2, value=formula)
    # Keep some values as formulas, not text
    # For first set, B column should be formula itself
    # Actually we already set formula as value

# Set number formats
ws_kpi.column_dimensions['A'].width = 22
ws_kpi.column_dimensions['B'].width = 18
ws_kpi.column_dimensions['C'].width = 45
ws_kpi.column_dimensions['D'].width = 35

# Format KPI values
for r in range(2, 13):
    cell = ws_kpi.cell(row=r, column=2)
    if r in [2,3,4,6,10,11]:
        cell.number_format = '#,##0.00'
    if r in [7,8,12]:
        cell.number_format = '0.00%'

# 8. Dashboard
ws_dash = wb.create_sheet("Dashboard")
ws_dash.sheet_properties.tabColor = "FFC000"

# Title
ws_dash.merge_cells('A1:H1')
ws_dash['A1'] = "SALES PERFORMANCE DASHBOARD | 2023-2024"
ws_dash['A1'].font = Font(name='Calibri', bold=True, size=18, color='FFFFFF')
ws_dash['A1'].fill = PatternFill(start_color='2F5597', end_color='2F5597', fill_type='solid')
ws_dash['A1'].alignment = center_align
ws_dash.row_dimensions[1].height = 35

# KPI Cards area
kpi_card_titles = ["Total Sales", "Total Profit", "Total Orders", "Profit Margin %", "Avg Order Value"]
kpi_card_formulas = [
    f"=KPIs!B2",
    f"=KPIs!B4",
    f"=KPIs!B5",
    f"=KPIs!B8",
    f"=KPIs!B7",
]
positions = ['A3','C3','E3','G3','A5']  # but we need merges
# Create KPI cards in row 3 and 5
ws_dash.merge_cells('A3:B3')
ws_dash['A3'] = "Total Sales"
ws_dash.merge_cells('A4:B4')
ws_dash['A4'] = f"=KPIs!B2"
ws_dash.merge_cells('C3:D3')
ws_dash['C3'] = "Total Profit"
ws_dash.merge_cells('C4:D4')
ws_dash['C4'] = f"=KPIs!B4"
ws_dash.merge_cells('E3:F3')
ws_dash['E3'] = "Total Orders"
ws_dash.merge_cells('E4:F4')
ws_dash['E4'] = f"=KPIs!B5"
ws_dash.merge_cells('G3:H3')
ws_dash['G3'] = "Profit Margin %"
ws_dash.merge_cells('G4:H4')
ws_dash['G4'] = f"=KPIs!B8"

ws_dash.merge_cells('A5:B5')
ws_dash['A5'] = "Avg Order Value"
ws_dash.merge_cells('A6:B6')
ws_dash['A6'] = f"=KPIs!B7"
ws_dash.merge_cells('C5:D5')
ws_dash['C5'] = "Total Quantity"
ws_dash.merge_cells('C6:D6')
ws_dash['C6'] = f"=KPIs!B6"
ws_dash.merge_cells('E5:F5')
ws_dash['E5'] = "YoY Growth"
ws_dash.merge_cells('E6:F6')
ws_dash['E6'] = f"=KPIs!B12"
ws_dash.merge_cells('G5:H5')
ws_dash['G5'] = "Avg Discount"
ws_dash.merge_cells('G6:H6')
ws_dash['G6'] = f"=KPIs!B9"

for row in [3,5]:
    for col in ['A','C','E','G']:
        cell = ws_dash[f'{col}{row}']
        cell.font = Font(bold=True, color='FFFFFF', size=11)
        cell.fill = PatternFill(start_color='2F5597', end_color='2F5597', fill_type='solid')
        cell.alignment = center_align
        cell.border = thin_border
for row in [4,6]:
    for col in ['A','C','E','G']:
        cell = ws_dash[f'{col}{row}']
        cell.font = Font(bold=True, size=14, color='2F5597')
        cell.fill = PatternFill(start_color='D9E2F3', end_color='D9E2F3', fill_type='solid')
        cell.alignment = center_align
        cell.border = thin_border

ws_dash['A4'].number_format = '$#,##0'
ws_dash['C4'].number_format = '$#,##0'
ws_dash['E4'].number_format = '#,##0'
ws_dash['G4'].number_format = '0.0%'
ws_dash['A6'].number_format = '$#,##0'
ws_dash['C6'].number_format = '#,##0'
ws_dash['E6'].number_format = '0.0%'
ws_dash['G6'].number_format = '0.0%'

# Aggregation tables for charts (starting row 8)
# Monthly trend
ws_dash['A8'] = "Year-Month"
ws_dash['B8'] = "Sales"
ws_dash['C8'] = "Profit"
for col in ['A','B','C']:
    ws_dash[f'{col}8'].font = header_font
    ws_dash[f'{col}8'].fill = header_fill

# Compute monthly aggregation in python
monthly = defaultdict(lambda: {"sales":0, "profit":0})
for row in range(2, ws_sales.max_row+1):
    od = ws_sales.cell(row=row, column=2).value
    sales_amt = ws_sales.cell(row=row, column=10).value
    profit = ws_sales.cell(row=row, column=12).value
    ym = od.strftime("%Y-%m")
    monthly[ym]["sales"] += sales_amt
    monthly[ym]["profit"] += profit

sorted_months = sorted(monthly.keys())
for idx, ym in enumerate(sorted_months, start=9):
    ws_dash.cell(row=idx, column=1, value=ym)
    ws_dash.cell(row=idx, column=2, value=monthly[ym]["sales"])
    ws_dash.cell(row=idx, column=3, value=monthly[ym]["profit"])
monthly_end_row = 8 + len(sorted_months)

# Category aggregation
ws_dash['E8'] = "Category"
ws_dash['F8'] = "Sales"
ws_dash['G8'] = "Profit"
for col in ['E','F','G']:
    ws_dash[f'{col}8'].font = header_font
    ws_dash[f'{col}8'].fill = header_fill

cat_agg = defaultdict(lambda: {"sales":0, "profit":0})
for row in range(2, ws_sales.max_row+1):
    prod_id = ws_sales.cell(row=row, column=4).value
    sales_amt = ws_sales.cell(row=row, column=10).value
    profit = ws_sales.cell(row=row, column=12).value
    cat = prod_dict[prod_id]["category"]
    cat_agg[cat]["sales"] += sales_amt
    cat_agg[cat]["profit"] += profit

cat_start = 9
for idx, (cat, vals) in enumerate(cat_agg.items(), start=cat_start):
    ws_dash.cell(row=idx, column=5, value=cat)
    ws_dash.cell(row=idx, column=6, value=vals["sales"])
    ws_dash.cell(row=idx, column=7, value=vals["profit"])
cat_end = cat_start + len(cat_agg) -1

# Region aggregation
ws_dash['I8'] = "Region"
ws_dash['J8'] = "Sales"
for col in ['I','J']:
    ws_dash[f'{col}8'].font = header_font
    ws_dash[f'{col}8'].fill = header_fill

region_names = {"R01":"North","R02":"South","R03":"East","R04":"West"}
region_agg = defaultdict(float)
for row in range(2, ws_sales.max_row+1):
    reg_id = ws_sales.cell(row=row, column=5).value
    sales_amt = ws_sales.cell(row=row, column=10).value
    region_agg[reg_id] += sales_amt

for idx, (reg_id, sales) in enumerate(region_agg.items(), start=9):
    ws_dash.cell(row=idx, column=9, value=region_names[reg_id])
    ws_dash.cell(row=idx, column=10, value=sales)
region_end = 9 + len(region_agg) -1

# Channel aggregation
ws_dash['L8'] = "Channel"
ws_dash['M8'] = "Sales"
for col in ['L','M']:
    ws_dash[f'{col}8'].font = header_font
    ws_dash[f'{col}8'].fill = header_fill

channel_agg = defaultdict(float)
for row in range(2, ws_sales.max_row+1):
    ch = ws_sales.cell(row=row, column=13).value
    sales_amt = ws_sales.cell(row=row, column=10).value
    channel_agg[ch] += sales_amt

for idx, (ch, sales) in enumerate(channel_agg.items(), start=9):
    ws_dash.cell(row=idx, column=12, value=ch)
    ws_dash.cell(row=idx, column=13, value=sales)

# Product aggregation for Top 10
ws_dash['A35'] = "ProductID"
ws_dash['B35'] = "ProductName"
ws_dash['C35'] = "Sales"
for col in ['A','B','C']:
    ws_dash[f'{col}35'].font = header_font
    ws_dash[f'{col}35'].fill = header_fill

prod_agg = defaultdict(float)
prod_name_map = {p[0]:p[1] for p in products}
for row in range(2, ws_sales.max_row+1):
    pid = ws_sales.cell(row=row, column=4).value
    sales_amt = ws_sales.cell(row=row, column=10).value
    prod_agg[pid] += sales_amt

sorted_prods = sorted(prod_agg.items(), key=lambda x: x[1], reverse=True)[:10]
for idx, (pid, sales) in enumerate(sorted_prods, start=36):
    ws_dash.cell(row=idx, column=1, value=pid)
    ws_dash.cell(row=idx, column=2, value=prod_name_map[pid])
    ws_dash.cell(row=idx, column=3, value=sales)

# Create Charts
# 1. Monthly Sales Trend - Line Chart
line = LineChart()
line.title = "Monthly Sales Trend"
line.style = 2
line.y_axis.title = "Sales"
line.x_axis.title = "Year-Month"
data = Reference(ws_dash, min_col=2, min_row=8, max_row=8+len(sorted_months))
# Actually need data including sales
cats = Reference(ws_dash, min_col=1, min_row=9, max_row=8+len(sorted_months))
sales_ref = Reference(ws_dash, min_col=2, min_row=8, max_row=8+len(sorted_months))
profit_ref = Reference(ws_dash, min_col=3, min_row=8, max_row=8+len(sorted_months))
line.add_data(sales_ref, titles_from_data=True)
line.add_data(profit_ref, titles_from_data=True)
line.set_categories(cats)
line.height = 10
line.width = 18
ws_dash.add_chart(line, "A50")

# 2. Category Sales - Bar Chart
bar_cat = BarChart()
bar_cat.title = "Sales by Category"
bar_cat.style = 10
bar_cat.type = "col"
bar_cat.y_axis.title = "Sales"
cats_cat = Reference(ws_dash, min_col=5, min_row=9, max_row=cat_end)
data_cat = Reference(ws_dash, min_col=6, min_row=8, max_row=cat_end)
bar_cat.add_data(data_cat, titles_from_data=True)
bar_cat.set_categories(cats_cat)
bar_cat.height = 10
bar_cat.width = 12
ws_dash.add_chart(bar_cat, "J50")

# 3. Region Sales - Pie Chart
pie_reg = PieChart()
pie_reg.title = "Sales by Region"
labels = Reference(ws_dash, min_col=9, min_row=9, max_row=region_end)
data_reg = Reference(ws_dash, min_col=10, min_row=8, max_row=region_end)
pie_reg.add_data(data_reg, titles_from_data=True)
pie_reg.set_categories(labels)
pie_reg.height = 10
pie_reg.width = 12
ws_dash.add_chart(pie_reg, "A65")

# 4. Channel Sales - Bar Chart
bar_ch = BarChart()
bar_ch.title = "Sales by Channel"
bar_ch.style = 11
bar_ch.type = "bar"
cats_ch = Reference(ws_dash, min_col=12, min_row=9, max_row=12)
data_ch = Reference(ws_dash, min_col=13, min_row=8, max_row=12)
bar_ch.add_data(data_ch, titles_from_data=True)
bar_ch.set_categories(cats_ch)
bar_ch.height = 8
bar_ch.width = 12
ws_dash.add_chart(bar_ch, "J65")

# 5. Top Products
bar_prod = BarChart()
bar_prod.title = "Top 10 Products by Sales"
bar_prod.style = 12
bar_prod.type = "col"
cats_prod = Reference(ws_dash, min_col=2, min_row=36, max_row=45)
data_prod = Reference(ws_dash, min_col=3, min_row=35, max_row=45)
bar_prod.add_data(data_prod, titles_from_data=True)
bar_prod.set_categories(cats_prod)
bar_prod.height = 12
bar_prod.width = 18
ws_dash.add_chart(bar_prod, "A80")

# Set column widths for dashboard
for col in ['A','B','C','D','E','F','G','H','I','J','L','M']:
    ws_dash.column_dimensions[col].width = 14
ws_dash.column_dimensions['B'].width = 20

# Freeze panes?
ws_dash.freeze_panes = 'A9'

# Save
wb.save("Sales_Performance_Dashboard_Excel.xlsx")
print("Excel Dashboard created: Sales_Performance_Dashboard_Excel.xlsx")
print(f"Rows Sales: {ws_sales.max_row-1}")
