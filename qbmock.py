import pandas as pd
import numpy as np
import random
from datetime import datetime, timedelta

# Helper function to generate random dates
def generate_dates(start_date, end_date, n):
    dates = pd.date_range(start=start_date, end=end_date).to_pydatetime().tolist()
    return random.choices(dates, k=n)

# Generate Expense Report Data
def generate_expense_report():
    categories = ["Food Supplies", "Beverages", "Utilities", "Cleaning Supplies", "Rent", "Staff Salaries", "Marketing"]
    vendors = ["Vendor A", "Vendor B", "Vendor C", "Vendor D", "Vendor E"]
    products = ["Beer", "Whiskey", "Soda", "Napkins", "Cleaning Spray", "Electricity", "Gas"]
    
    data = {
        "Date": generate_dates("2018-01-01", "2023-12-31", 1000),
        "Expense_Category": [random.choice(categories) for _ in range(1000)],
        "Vendor": [random.choice(vendors) for _ in range(1000)],
        "Product": [random.choice(products) for _ in range(1000)],
        "Amount": [round(random.uniform(50, 2000), 2) for _ in range(1000)]
    }
    return pd.DataFrame(data)

# Generate Cash Flow Statement Data
def generate_cash_flow_statement():
    dates = pd.date_range(start="2018-01-01", end="2023-12-31", freq='M')
    data = {
        "Date": dates,
        "Cash_Inflow": [round(random.uniform(20000, 50000), 2) for _ in range(len(dates))],
        "Cash_Outflow": [round(random.uniform(15000, 45000), 2) for _ in range(len(dates))]
    }
    return pd.DataFrame(data)

# Generate Balance Sheet Data
def generate_balance_sheet():
    dates = pd.date_range(start="2018-01-01", end="2023-12-31", freq='M')
    data = {
        "Date": dates,
        "Current_Assets": [round(random.uniform(50000, 100000), 2) for _ in range(len(dates))],
        "Current_Liabilities": [round(random.uniform(20000, 70000), 2) for _ in range(len(dates))],
        "Long_Term_Debt": [round(random.uniform(50000, 150000), 2) for _ in range(len(dates))],
        "Equity": [round(random.uniform(100000, 300000), 2) for _ in range(len(dates))]
    }
    return pd.DataFrame(data)

# Generate Profit & Loss Statement Data
def generate_profit_loss_statement():
    dates = pd.date_range(start="2018-01-01", end="2023-12-31", freq='M')
    data = {
        "Date": dates,
        "Revenue": [round(random.uniform(50000, 150000), 2) for _ in range(len(dates))],
        "COGS": [round(random.uniform(20000, 50000), 2) for _ in range(len(dates))],
        "Operating_Expenses": [round(random.uniform(10000, 40000), 2) for _ in range(len(dates))],
        "Net_Income": lambda x: x["Revenue"] - x["COGS"] - x["Operating_Expenses"]
    }
    df = pd.DataFrame(data)
    df["Net_Income"] = df["Revenue"] - df["COGS"] - df["Operating_Expenses"]
    return df

# Write data to Excel
with pd.ExcelWriter("Sports_Bar_QuickBooks_Dummy_Data_5_Years.xlsx") as writer:
    generate_expense_report().to_excel(writer, sheet_name="Expense Report", index=False)
    generate_cash_flow_statement().to_excel(writer, sheet_name="Cash Flow Statement", index=False)
    generate_balance_sheet().to_excel(writer, sheet_name="Balance Sheet", index=False)
    generate_profit_loss_statement().to_excel(writer, sheet_name="Profit & Loss Statement", index=False)

print("Excel file created: Sports_Bar_QuickBooks_Dummy_Data_5_Years.xlsx")
