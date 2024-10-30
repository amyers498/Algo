import pandas as pd
import numpy as np

# Load Financial Data from Sports Bar QuickBooks file
def load_quickbooks_data(file_path):
    print("Loading QuickBooks data...")
    expense_df = pd.read_excel(file_path, sheet_name="Expense Report")
    cash_flow_df = pd.read_excel(file_path, sheet_name="Cash Flow Statement")
    balance_sheet_df = pd.read_excel(file_path, sheet_name="Balance Sheet")
    pnl_df = pd.read_excel(file_path, sheet_name="Profit & Loss Statement")
    print("Data loaded successfully!")
    return expense_df, cash_flow_df, balance_sheet_df, pnl_df

# Identify Primary Commodity Cost
def select_primary_cost(expense_df):
    print("Selecting primary cost driver...")
    expense_summary = expense_df.groupby("Expense_Category")["Amount"].sum().sort_values(ascending=False)
    primary_cost_category = expense_summary.idxmax()
    primary_cost = expense_summary.max()
    print(f"Primary Cost Category: {primary_cost_category}, Primary Cost: {primary_cost}")
    return primary_cost_category, primary_cost

# Map Expense Category to Commodity
def map_expense_to_commodity(expense_category):
    commodity_mapping = {
        "Food Supplies": "Corn",
        "Beverages": "Sugar",
        "Utilities": "Natural Gas",
        "Cleaning Supplies": "Cotton",
        "Packaging": "Wheat"
    }
    return commodity_mapping.get(expense_category, "Corn")  # Default to Corn if no match

# Load Commodity Data
def load_commodity_data(file_path):
    print("Loading commodity data...")
    commodity_df = pd.read_excel(file_path)
    print("Commodity data loaded successfully!")
    return commodity_df

# Calculate Volatility and Seasonality for Commodity
def calculate_volatility_and_seasonality(commodity_df):
    print("Calculating volatility and seasonality...")
    commodity_df["Daily_Return"] = commodity_df["Commodity_Price"].pct_change()
    volatility = commodity_df["Daily_Return"].std() * np.sqrt(252)
    commodity_df["Month"] = pd.to_datetime(commodity_df["Date"]).dt.month
    monthly_avg = commodity_df.groupby("Month")["Commodity_Price"].mean()
    high_season = monthly_avg.idxmax() if monthly_avg.max() > 1.5 * monthly_avg.mean() else None
    print(f"Volatility: {volatility * 100:.2f}%, High Season: {high_season}")
    return volatility * 100, high_season

# Assess Financial Profile
def assess_financial_profile(cash_flow_df, balance_sheet_df, pnl_df):
    print("Assessing financial profile...")
    avg_cash_inflow = cash_flow_df["Cash_Inflow"].mean()
    avg_cash_outflow = cash_flow_df["Cash_Outflow"].mean()
    cash_flow_stability = avg_cash_inflow - avg_cash_outflow
    
    avg_assets = balance_sheet_df["Current_Assets"].mean()
    avg_liabilities = balance_sheet_df["Current_Liabilities"].mean()
    current_ratio = avg_assets / avg_liabilities
    
    avg_long_term_debt = balance_sheet_df["Long_Term_Debt"].mean()
    avg_equity = balance_sheet_df["Equity"].mean()
    debt_to_equity_ratio = (avg_liabilities + avg_long_term_debt) / avg_equity
    
    avg_revenue = pnl_df["Revenue"].mean()
    avg_net_income = pnl_df["Net_Income"].mean()
    profit_margin = avg_net_income / avg_revenue
    
    risk_tolerance = "High" if debt_to_equity_ratio > 2 or current_ratio < 1.5 or profit_margin < 0.05 else "Moderate"
    
    financial_profile = {
        "cash_flow_stability": cash_flow_stability,
        "current_ratio": current_ratio,
        "debt_to_equity_ratio": debt_to_equity_ratio,
        "profit_margin": profit_margin,
        "risk_tolerance": risk_tolerance
    }
    print("Financial Profile:", financial_profile)
    return financial_profile

# Select Contract Durations based on Hedge Duration and Risk Tolerance
def select_contract_durations(hedge_duration):
    contract_durations = {
        "Short-Term (1-3 months)": [1, 2, 3],
        "Medium-Term (3-12 months)": [3, 6, 12],
        "Long-Term (12+ months)": [12, 18, 24]
    }
    return contract_durations[hedge_duration]

# Recommend Hedge Duration
def recommend_hedge_duration(volatility, financial_profile, expense_df):
    print("Recommending hedge duration...")
    monthly_expenses = expense_df.groupby(expense_df['Date'].dt.month)["Amount"].sum()
    high_season = monthly_expenses.idxmax() if monthly_expenses.max() > 1.5 * monthly_expenses.mean() else None
    
    if volatility > 25 or financial_profile["risk_tolerance"] == "High" or high_season:
        hedge_duration = "Short-Term (1-3 months)"
    elif 15 < volatility <= 25 or financial_profile["risk_tolerance"] == "Moderate":
        hedge_duration = "Medium-Term (3-12 months)"
    else:
        hedge_duration = "Long-Term (12+ months)"
    
    print(f"Recommended Hedge Duration: {hedge_duration}")
    return hedge_duration

# Recommend and Select Best Contracts
def recommend_contracts_and_select_best(commodity_df, primary_commodity, hedge_duration, financial_profile):
    print("Analyzing and recommending contracts...")
    commodity_data = commodity_df[commodity_df['Commodity'] == primary_commodity]
    selected_durations = select_contract_durations(hedge_duration)
    contracts = []

    for duration in selected_durations:
        contract_info = {"Duration (months)": duration}
        historical_data = commodity_data.tail(duration * 20)
        if historical_data.empty or len(historical_data) < 2:
            print(f"Not enough data for {duration}-month duration. Skipping this contract.")
            continue

        historical_data['Daily_Return'] = historical_data['Commodity_Price'].pct_change()
        historical_return = historical_data['Daily_Return'].mean() * duration * 20
        historical_volatility = historical_data['Daily_Return'].std() * np.sqrt(252)
        
        contract_info.update({
            "Historical Average Return (%)": round(historical_return * 100, 2),
            "Historical Volatility (%)": round(historical_volatility * 100, 2)
        })

        try:
            recent_trend = historical_data['Commodity_Price'].iloc[-1] / historical_data['Commodity_Price'].iloc[0]
            expected_future_return = historical_return * recent_trend
            expected_future_volatility = historical_volatility * np.sqrt(recent_trend)
            contract_info.update({
                "Expected Future Return (%)": round(expected_future_return * 100, 2),
                "Expected Future Volatility (%)": round(expected_future_volatility * 100, 2)
            })
        except IndexError:
            print(f"Insufficient data points for forward analysis on {duration}-month duration.")
            continue

        contracts.append(contract_info)

    if contracts:
        if financial_profile["risk_tolerance"] == "High":
            best_contract = min(contracts, key=lambda x: x["Historical Volatility (%)"])
        else:
            best_contract = max(contracts, key=lambda x: x["Expected Future Return (%)"] / (x["Expected Future Volatility (%)"] + 1))
        print("Best Contract:", best_contract)
        return best_contract
    else:
        print("No viable contracts found.")
        return None

# Main Execution
quickbooks_path = "Sports_Bar_QuickBooks_Dummy_Data_24_Months.xlsx"
commodity_data_path = "Realistic_Ag_Options_Data_Last_3_Years.xlsx"

# Load data
expense_df, cash_flow_df, balance_sheet_df, pnl_df = load_quickbooks_data(quickbooks_path)
primary_cost_category, primary_cost = select_primary_cost(expense_df)
mapped_commodity = map_expense_to_commodity(primary_cost_category)
print(f"Mapped Commodity for '{primary_cost_category}': {mapped_commodity}")

commodity_df = load_commodity_data(commodity_data_path)

# Calculate volatility and seasonality
volatility, high_season = calculate_volatility_and_seasonality(commodity_df[commodity_df["Commodity"] == mapped_commodity])

# Assess financial profile
financial_profile = assess_financial_profile(cash_flow_df, balance_sheet_df, pnl_df)

# Recommend hedge duration
hedge_duration = recommend_hedge_duration(volatility, financial_profile, expense_df)

# Recommend and select the best contract
best_contract = recommend_contracts_and_select_best(commodity_df, mapped_commodity, hedge_duration, financial_profile)

# Final Output
print("\nFinal Output:")
print({
    "Primary Cost Category": primary_cost_category,
    "Primary Cost": primary_cost,
    "Mapped Commodity": mapped_commodity,
    "Financial Profile": financial_profile,
    "Commodity Volatility (%)": volatility,
    "High Season": high_season,
        "Recommended Hedge Duration": hedge_duration,
    "Best Contract": best_contract  # Display the best contract recommendation
})

