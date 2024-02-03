"""
Monthly Sales Report Automator
Client: ShopFlow (small e-commerce brand)
Delivered: February 2024
"""

import pandas as pd
import glob
import os

# ---------------------------
# 1. Load all weekly files
# ---------------------------
path = "weekly_sales"
all_files = glob.glob(os.path.join(path, "*.xlsx"))
if not all_files:
    raise FileNotFoundError("No weekly sales files found.")

df_list = []
for file in all_files:
    df = pd.read_excel(file, dtype={'Quantity': 'object', 'Price': 'object'})
    # Drop completely blank rows
    df.dropna(how='all', inplace=True)
    df_list.append(df)

df = pd.concat(df_list, ignore_index=True)

# ---------------------------
# 2. Clean data
# ---------------------------

# 2.1 Remove rows where both Quantity and Price are missing
df.dropna(subset=['Quantity', 'Price'], how='all', inplace=True)

# 2.2 Convert Quantity to numeric, forcing errors to NaN, then drop NaN in Quantity
df['Quantity'] = pd.to_numeric(df['Quantity'], errors='coerce')
df.dropna(subset=['Quantity'], inplace=True)

# 2.3 Convert Price to numeric, fill missing with median later
df['Price'] = pd.to_numeric(df['Price'], errors='coerce')

# 2.4 Fix Product names: strip whitespace, title case, remove extra spaces
df['Product'] = df['Product'].astype(str).str.strip().str.title()
# Manually merge known typos if any remain (e.g., 'Widget A ' becomes 'Widget A' after strip)
# The title() will change 'widget a' -> 'Widget A'

# 2.5 Remove exact duplicate rows
df.drop_duplicates(inplace=True)

# 2.6 Handle negative quantities (refunds) — flag them for client but keep
df['Is_Refund'] = df['Quantity'] < 0

# 2.7 Parse dates with mixed formats
df['Date'] = pd.to_datetime(df['Date'], dayfirst=False, errors='coerce')
df.dropna(subset=['Date'], inplace=True)

# 2.8 Fill missing Price with median price of that product; if still missing, global median
df['Price'] = df.groupby('Product')['Price'].transform(lambda x: x.fillna(x.median()))
overall_median = df['Price'].median()
df['Price'] = df['Price'].fillna(overall_median)

# Now make Quantity positive for revenue calculation (refunds will have negative revenue)
df['Revenue'] = df['Quantity'] * df['Price']
df['Week'] = df['Date'].dt.isocalendar().week

# ---------------------------
# 3. Summaries
# ---------------------------
product_summary = df.groupby('Product').agg(
    Total_Revenue=('Revenue', 'sum'),
    Units_Sold=('Quantity', lambda x: x[x>0].sum()),
    Refunds=('Is_Refund', 'sum')
).sort_values('Total_Revenue', ascending=False)

weekly_summary = df.groupby('Week').agg(
    Revenue=('Revenue', 'sum'),
    Orders=('Quantity', lambda x: (x>0).sum()),
    Refunds=('Is_Refund', 'sum')
).reset_index()

top5 = product_summary.head(5)

# ---------------------------
# 4. Write to Excel with charts
# ---------------------------
with pd.ExcelWriter('Monthly_Report.xlsx', engine='xlsxwriter') as writer:
    df.to_excel(writer, sheet_name='Cleaned Data', index=False)
    product_summary.to_excel(writer, sheet_name='Product Summary')
    weekly_summary.to_excel(writer, sheet_name='Weekly Trend')
    top5.to_excel(writer, sheet_name='Top 5 Products')

    workbook = writer.book

    # Chart 1: Top 5 Products bar chart
    chart1 = workbook.add_chart({'type': 'bar'})
    chart1.add_series({
        'name': 'Revenue',
        'categories': f'=Top 5 Products!$A$2:$A$6',
        'values': f'=Top 5 Products!$B$2:$B$6',
    })
    chart1.set_title({'name': 'Top 5 Products by Revenue'})
    chart1.set_x_axis({'name': 'Revenue (USD)'})
    writer.sheets['Top 5 Products'].insert_chart('D2', chart1)

    # Chart 2: Weekly Revenue Trend line chart
    chart2 = workbook.add_chart({'type': 'line'})
    chart2.add_series({
        'name': 'Weekly Revenue',
        'categories': f'=Weekly Trend!$A$2:$A$5',
        'values': f'=Weekly Trend!$B$2:$B$5',
    })
    chart2.set_title({'name': 'Weekly Revenue Trend'})
    chart2.set_y_axis({'name': 'Revenue (USD)'})
    writer.sheets['Weekly Trend'].insert_chart('D2', chart2)

print("Monthly Report saved as 'Monthly_Report.xlsx'")