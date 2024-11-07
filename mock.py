import pandas as pd
import numpy as np

# Create date range for the mock data (monthly over 5 years)
dates = pd.date_range(start="2019-01-01", end="2023-12-31", freq="M")

# Define expense amounts for categories, creating correlations with commodity prices
# Sugar expenses will correlate with sugar commodity prices
# Cooking oil expenses will correlate with diesel commodity prices
sugar_expenses = 100 + 10 * np.sin(2 * np.pi * dates.month / 12)  # Stronger seasonal pattern for correlation
cooking_oil_expenses = 150 + 12 * np.sin(2 * np.pi * dates.month / 12)
labor_expenses = 300 + np.random.normal(0, 20, len(dates))  # No correlation to commodity data

# Expense Report
expense_report = pd.DataFrame({
    "Date": np.tile(dates, 3),
    "Expense_Category": ["Sugar"] * len(dates) + ["Cooking Oil"] * len(dates) + ["Labor"] * len(dates),
    "Amount": np.concatenate([sugar_expenses, cooking_oil_expenses, labor_expenses])
})

# Cash Flow Statement
cash_inflow = 5000 + 500 * np.random.normal(0, 1, len(dates))
cash_outflow = 4800 + 400 * np.random.normal(0, 1, len(dates))
cash_flow_statement = pd.DataFrame({"Date": dates, "Cash_Inflow": cash_inflow, "Cash_Outflow": cash_outflow})

# Balance Sheet
balance_sheet = pd.DataFrame({
    "Date": dates,
    "Current_Assets": 10000 + 1000 * np.random.normal(0, 1, len(dates)),
    "Current_Liabilities": 8000 + 800 * np.random.normal(0, 1, len(dates)),
    "Long_Term_Debt": 5000 + 500 * np.random.normal(0, 1, len(dates)),
    "Equity": 7000 + 700 * np.random.normal(0, 1, len(dates))
})

# Profit & Loss Statement
revenue = 12000 + 1000 * np.random.normal(0, 1, len(dates))
net_income = revenue - (cash_outflow + labor_expenses)
pnl_statement = pd.DataFrame({"Date": dates, "Revenue": revenue, "Net_Income": net_income})

# Write QuickBooks data to Excel
with pd.ExcelWriter("/mnt/data/Mock_QuickBooks_Expenses_Correlated.xlsx") as writer:
    expense_report.to_excel(writer, sheet_name="Expense Report", index=False)
    cash_flow_statement.to_excel(writer, sheet_name="Cash Flow Statement", index=False)
    balance_sheet.to_excel(writer, sheet_name="Balance Sheet", index=False)
    pnl_statement.to_excel(writer, sheet_name="Profit & Loss Statement", index=False)

# Generate commodity data with similar patterns for correlation
commodity_dates = pd.date_range(start="2019-01-01", end="2023-12-31", freq="M")
sugar_prices = 50 + 10 * np.sin(2 * np.pi * commodity_dates.month / 12)  # Matches sugar expenses
diesel_prices = 70 + 12 * np.sin(2 * np.pi * commodity_dates.month / 12)  # Matches cooking oil expenses
wheat_prices = 90 + np.random.normal(0, 10, len(commodity_dates))  # Unrelated commodity

# Create commodity data
commodity_df = pd.DataFrame({
    "Date": np.tile(commodity_dates, 3),
    "Commodity": ["Sugar"] * len(commodity_dates) + ["Diesel"] * len(commodity_dates) + ["Wheat"] * len(commodity_dates),
    "Commodity_Price": np.concatenate([sugar_prices, diesel_prices, wheat_prices])
})

# Write commodity data to Excel
commodity_df.to_excel("/mnt/data/Mock_Commodity_Data_Correlated.xlsx", sheet_name="Commodity Data", index=False)
