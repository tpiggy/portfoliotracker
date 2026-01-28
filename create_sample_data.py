#!/usr/bin/env python3
"""Create a sample investment tracking ODS file with realistic portfolio data."""

import pyexcel

# Sample investment data organized by account
# Each row: Account, Ticker, Name, Asset Class, Shares, Price, Value
investments = [
    # 401(k) Account
    ["401(k)", "VTI", "Vanguard Total Stock Market ETF", "US Stocks", 150, 245.80, 36870.00],
    ["401(k)", "VXUS", "Vanguard Total International Stock ETF", "International Stocks", 100, 62.45, 6245.00],
    ["401(k)", "BND", "Vanguard Total Bond Market ETF", "Bonds", 200, 72.30, 14460.00],
    ["401(k)", "VNQ", "Vanguard Real Estate ETF", "Real Estate", 50, 89.20, 4460.00],

    # Roth IRA
    ["Roth IRA", "VOO", "Vanguard S&P 500 ETF", "US Stocks", 80, 478.50, 38280.00],
    ["Roth IRA", "QQQ", "Invesco QQQ Trust", "US Stocks", 45, 485.20, 21834.00],
    ["Roth IRA", "SCHD", "Schwab US Dividend Equity ETF", "US Stocks", 120, 82.15, 9858.00],

    # Taxable Brokerage
    ["Taxable Brokerage", "AAPL", "Apple Inc.", "US Stocks", 100, 189.45, 18945.00],
    ["Taxable Brokerage", "MSFT", "Microsoft Corporation", "US Stocks", 75, 415.80, 31185.00],
    ["Taxable Brokerage", "GOOGL", "Alphabet Inc.", "US Stocks", 50, 175.20, 8760.00],
    ["Taxable Brokerage", "AMZN", "Amazon.com Inc.", "US Stocks", 60, 198.50, 11910.00],
    ["Taxable Brokerage", "NVDA", "NVIDIA Corporation", "US Stocks", 40, 875.30, 35012.00],
    ["Taxable Brokerage", "BRK.B", "Berkshire Hathaway Inc.", "US Stocks", 30, 458.70, 13761.00],

    # HSA Account
    ["HSA", "FXAIX", "Fidelity 500 Index Fund", "US Stocks", 85, 198.40, 16864.00],
    ["HSA", "FXNAX", "Fidelity US Bond Index Fund", "Bonds", 150, 10.85, 1627.50],

    # 529 Education Savings
    ["529 Plan", "VTTSX", "Vanguard Target Retirement 2060", "Target Date", 400, 48.90, 19560.00],

    # Emergency Fund (High-Yield Savings proxy as Money Market)
    ["Emergency Fund", "VMFXX", "Vanguard Federal Money Market", "Cash", 250, 100.00, 25000.00],

    # Crypto Account
    ["Crypto", "BTC", "Bitcoin", "Cryptocurrency", 0.5, 68500.00, 34250.00],
    ["Crypto", "ETH", "Ethereum", "Cryptocurrency", 5.2, 3850.00, 20020.00],
]

# Create the spreadsheet with headers
header = ["Account", "Ticker", "Name", "Asset Class", "Shares", "Price", "Value"]
data = [header] + investments

# Save as ODS file
pyexcel.save_as(array=data, dest_file_name="/home/user/portfoliotracker/investment_portfolio.ods")

print("Created: investment_portfolio.ods")
print(f"Total records: {len(investments)} holdings across multiple accounts")

# Calculate and display summary
total_value = sum(row[6] for row in investments)
print(f"Total Portfolio Value: ${total_value:,.2f}")
