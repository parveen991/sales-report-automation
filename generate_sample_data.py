import pandas as pd
import numpy as np
import os
from datetime import datetime, timedelta
import random

os.makedirs('weekly_sales', exist_ok=True)

products = [
    'Widget A', 'Widget B', 'Gadget X', 'Gadget Y',
    'SuperTool', 'MegaPack', 'Widget A ',  # trailing space (typo)
    'widget a'  # inconsistent casing
]
base_prices = {
    'Widget A': 12.50, 'Widget B': 8.99,
    'Gadget X': 45.00, 'Gadget Y': 32.00,
    'SuperTool': 25.50, 'MegaPack': 67.00,
    'Widget A ': 12.50, 'widget a': 12.50
}
num_weeks = 4
start_date = datetime(2024, 3, 4)  # First week of March 2024

for week in range(num_weeks):
    rows = []
    week_start = start_date + timedelta(weeks=week)
    for day_offset in range(5):  # Monday to Friday
        current_date = week_start + timedelta(days=day_offset)
        for product in products:
            quantity = np.random.poisson(lam=3)  # Usually 0-8 sales
            if quantity == 0 and random.random() < 0.3:
                quantity = np.nan  # 30% chance of missing when zero
            price = base_prices.get(product, 15.00)

            # Messiness: occasional missing price, date in wrong format, duplicates
            if random.random() < 0.05:
                price = None
            if random.random() < 0.05:
                quantity = -abs(quantity)  # Negative "refund" row

            date_str = current_date.strftime('%Y-%m-%d') if random.random() < 0.9 else current_date.strftime('%d/%m/%Y')

            rows.append([date_str, product, quantity, price])

            # Duplicate row sometimes
            if random.random() < 0.04:
                rows.append([date_str, product, quantity, price])

    # Create DataFrame, add a blank row at the end
    df = pd.DataFrame(rows, columns=['Date', 'Product', 'Quantity', 'Price'])
    df = pd.concat([df, pd.DataFrame([['', '', '', '']], columns=df.columns)], ignore_index=True)

    filename = f'weekly_sales/sales_week_{week+1}_{week_start.strftime("%b%d")}.xlsx'
    df.to_excel(filename, index=False)
    print(f"Created {filename} — {len(df)-1} rows + 1 blank.")

print("Messy sample data ready.")