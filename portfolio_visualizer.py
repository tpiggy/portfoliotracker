#!/usr/bin/env python3
"""
Investment Portfolio Visualizer

Reads an investment tracking ODS file and creates visualizations including:
- Total Portfolio Value
- Largest Account
- Top Holding
- Asset Mix
"""

import sys
import pyexcel
import pandas as pd
import matplotlib.pyplot as plt
from pathlib import Path


# Asset class mappings based on ticker/name patterns
ASSET_CLASS_RULES = {
    'CASH$': 'Cash',
    'MMA': 'Mining - Copper',
    'AMR': 'Mining - Coal',
    'CVE': 'Mining - Coal',
    'WM': 'Mining - Gold',
    'FAR': 'Mining - Services',
    'TOU': 'Energy - Oil & Gas',
    'LAND': 'Real Estate',
    'LULU': 'Consumer Discretionary',
    'GURU': 'Consumer Staples',
    'PZZA': 'Consumer Discretionary',
    'RKLB': 'Aerospace & Defense',
    'CRSP': 'Biotech',
    'BEAM': 'Biotech',
}


def infer_asset_class(ticker: str, name: str) -> str:
    """Infer asset class from ticker or company name."""
    ticker = str(ticker).upper()
    name = str(name).lower()

    # Check direct ticker mapping
    if ticker in ASSET_CLASS_RULES:
        return ASSET_CLASS_RULES[ticker]

    # Infer from name patterns
    if 'mining' in name or 'coal' in name or 'metallurgical' in name:
        return 'Mining'
    if 'oil' in name or 'gas' in name or 'energy' in name:
        return 'Energy'
    if 'cash' in name:
        return 'Cash'
    if 'therapeutics' in name or 'biotech' in name or 'crispr' in name:
        return 'Biotech'
    if 'land' in name or 'real estate' in name or 'reit' in name:
        return 'Real Estate'

    return 'Equities'


def load_portfolio_data(ods_file: str) -> pd.DataFrame:
    """Load investment data from an ODS file into a pandas DataFrame."""
    data = pyexcel.get_array(file_name=ods_file)

    if not data or len(data) < 2:
        raise ValueError("ODS file is empty or has no data rows")

    # Detect format: check if first row looks like headers or data
    first_row = data[0]

    # Check for Portfolio.ods format (data first, headers later)
    # In this format, row 15 contains headers like 'Name', 'Ticker', etc.
    is_portfolio_format = False
    for i, row in enumerate(data):
        if len(row) > 0 and row[0] == 'Name' and len(row) > 1 and row[1] == 'Ticker':
            is_portfolio_format = True
            header_row = i
            break

    if is_portfolio_format:
        # Portfolio.ods format: data rows come before header row
        header = ['Name', 'Ticker', 'Price', 'Shares', 'Value', 'Percentage', 'Change', 'Extra1', 'Extra2']
        rows = []
        for row in data[:header_row]:
            # Skip empty rows or rows without valid data
            if len(row) >= 5 and row[0] and row[4]:
                try:
                    value = float(row[4]) if row[4] else 0
                    if value > 0:
                        rows.append(row[:9] if len(row) >= 9 else row + [''] * (9 - len(row)))
                except (ValueError, TypeError):
                    continue

        df = pd.DataFrame(rows, columns=header)
    else:
        # Standard format: first row is header
        header = data[0]
        rows = data[1:]
        df = pd.DataFrame(rows, columns=header)

    # Ensure numeric columns are proper types
    for col in ['Value', 'Shares', 'Price', 'Percentage', 'Change']:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

    # Add Asset Class column if not present
    if 'Asset Class' not in df.columns:
        df['Asset Class'] = df.apply(lambda row: infer_asset_class(row.get('Ticker', ''), row.get('Name', '')), axis=1)

    # Add Account column if not present (default to 'Main Portfolio')
    if 'Account' not in df.columns:
        df['Account'] = 'Main Portfolio'

    return df


def calculate_portfolio_metrics(df: pd.DataFrame) -> dict:
    """Calculate key portfolio metrics."""
    metrics = {}

    # Total Portfolio Value
    metrics['total_value'] = df['Value'].sum()

    # Account breakdown
    account_values = df.groupby('Account')['Value'].sum().sort_values(ascending=False)
    metrics['account_values'] = account_values
    metrics['largest_account'] = account_values.index[0]
    metrics['largest_account_value'] = account_values.iloc[0]

    # Top holding
    top_holding_idx = df['Value'].idxmax()
    metrics['top_holding'] = {
        'ticker': df.loc[top_holding_idx, 'Ticker'],
        'name': df.loc[top_holding_idx, 'Name'],
        'value': df.loc[top_holding_idx, 'Value'],
        'account': df.loc[top_holding_idx, 'Account']
    }

    # Asset class breakdown
    asset_values = df.groupby('Asset Class')['Value'].sum().sort_values(ascending=False)
    metrics['asset_mix'] = asset_values

    # Holdings by value (top 10)
    holdings = df[['Ticker', 'Name', 'Value']].sort_values('Value', ascending=False).head(10)
    metrics['top_holdings'] = holdings

    return metrics


def create_visualizations(df: pd.DataFrame, metrics: dict, output_dir: str = "."):
    """Create and save portfolio visualizations."""
    output_path = Path(output_dir)
    output_path.mkdir(exist_ok=True)

    # Set up the style
    plt.style.use('seaborn-v0_8-whitegrid')
    colors = plt.cm.Set3.colors

    # Create a figure with 4 subplots (2x2)
    fig, axes = plt.subplots(2, 2, figsize=(14, 12))
    fig.suptitle('Investment Portfolio Dashboard', fontsize=16, fontweight='bold', y=1.02)

    # 1. Total Portfolio Value - Summary Card
    ax1 = axes[0, 0]
    ax1.set_xlim(0, 10)
    ax1.set_ylim(0, 10)
    ax1.axis('off')

    # Title box
    ax1.text(5, 8.5, 'Total Portfolio Value', fontsize=14, fontweight='bold',
             ha='center', va='center')
    ax1.text(5, 6.5, f'${metrics["total_value"]:,.2f}', fontsize=28, fontweight='bold',
             ha='center', va='center', color='#2E7D32')

    # Additional summary stats
    num_accounts = len(metrics['account_values'])
    num_holdings = len(df)
    num_asset_classes = len(metrics['asset_mix'])

    if num_accounts <= 1:
        summary_text = f"{num_holdings} Holdings  |  {num_asset_classes} Sectors"
    else:
        summary_text = f"{num_holdings} Holdings  |  {num_accounts} Accounts  |  {num_asset_classes} Asset Classes"
    ax1.text(5, 4, summary_text, fontsize=11, ha='center', va='center', color='#666666')

    # Top holding info
    top = metrics['top_holding']
    ax1.text(5, 2, f"Top Holding: {top['ticker']} ({top['name']})",
             fontsize=10, ha='center', va='center', color='#1565C0')
    ax1.text(5, 1, f"${top['value']:,.2f} in {top['account']}",
             fontsize=10, ha='center', va='center', color='#666666')

    # 2. Sector/Holdings Breakdown - Horizontal Bar Chart
    ax2 = axes[0, 1]
    # If only one account, show asset class breakdown instead
    if len(metrics['account_values']) <= 1:
        breakdown_data = metrics['asset_mix']
        title = 'Value by Sector'
    else:
        breakdown_data = metrics['account_values']
        title = 'Value by Account'

    y_pos = range(len(breakdown_data))
    bars = ax2.barh(y_pos, breakdown_data.values, color=[colors[i % len(colors)] for i in range(len(breakdown_data))])
    ax2.set_yticks(y_pos)
    ax2.set_yticklabels(breakdown_data.index)
    ax2.set_xlabel('Value ($)')
    ax2.set_title(title, fontsize=12, fontweight='bold', pad=10)

    # Add value labels on bars
    for i, (bar, val) in enumerate(zip(bars, breakdown_data.values)):
        ax2.text(val + metrics['total_value'] * 0.01, bar.get_y() + bar.get_height()/2,
                 f'${val:,.0f}', va='center', fontsize=9)

    # Highlight largest item
    bars[0].set_color('#2E7D32')
    bars[0].set_edgecolor('#1B5E20')
    bars[0].set_linewidth(2)

    ax2.invert_yaxis()

    # 3. Asset Mix - Pie Chart
    ax3 = axes[1, 0]
    asset_data = metrics['asset_mix']

    # Create pie chart with percentages
    wedges, texts, autotexts = ax3.pie(
        asset_data.values,
        labels=asset_data.index,
        autopct=lambda pct: f'{pct:.1f}%' if pct > 3 else '',
        colors=[colors[i % len(colors)] for i in range(len(asset_data))],
        explode=[0.02] * len(asset_data),
        startangle=90
    )
    ax3.set_title('Asset Mix', fontsize=12, fontweight='bold', pad=10)

    # Style the percentage labels
    for autotext in autotexts:
        autotext.set_fontsize(9)
        autotext.set_fontweight('bold')

    # 4. Top 10 Holdings - Horizontal Bar Chart
    ax4 = axes[1, 1]
    top_holdings = metrics['top_holdings']
    y_pos = range(len(top_holdings))

    # Create labels with ticker and name
    labels = [f"{row['Ticker']}" for _, row in top_holdings.iterrows()]

    bars = ax4.barh(y_pos, top_holdings['Value'].values,
                    color=[colors[i % len(colors)] for i in range(len(top_holdings))])
    ax4.set_yticks(y_pos)
    ax4.set_yticklabels(labels)
    ax4.set_xlabel('Value ($)')
    ax4.set_title('Top 10 Holdings', fontsize=12, fontweight='bold', pad=10)

    # Add value labels
    for bar, val in zip(bars, top_holdings['Value'].values):
        ax4.text(val + metrics['total_value'] * 0.005, bar.get_y() + bar.get_height()/2,
                 f'${val:,.0f}', va='center', fontsize=9)

    # Highlight top holding
    bars[0].set_color('#1565C0')
    bars[0].set_edgecolor('#0D47A1')
    bars[0].set_linewidth(2)

    ax4.invert_yaxis()

    plt.tight_layout()

    # Save the combined dashboard
    dashboard_path = output_path / 'portfolio_dashboard.png'
    plt.savefig(dashboard_path, dpi=150, bbox_inches='tight', facecolor='white')
    print(f"Saved: {dashboard_path}")

    plt.close()

    return str(dashboard_path)


def print_summary(metrics: dict):
    """Print a text summary of the portfolio."""
    print("\n" + "=" * 60)
    print("INVESTMENT PORTFOLIO SUMMARY")
    print("=" * 60)

    print(f"\n{'Total Portfolio Value:':<25} ${metrics['total_value']:>15,.2f}")

    print(f"\n{'Largest Account:':<25} {metrics['largest_account']}")
    print(f"{'   Value:':<25} ${metrics['largest_account_value']:>15,.2f}")

    top = metrics['top_holding']
    print(f"\n{'Top Holding:':<25} {top['ticker']} - {top['name']}")
    print(f"{'   Value:':<25} ${top['value']:>15,.2f}")
    print(f"{'   Account:':<25} {top['account']}")

    print("\nAsset Mix:")
    print("-" * 40)
    for asset_class, value in metrics['asset_mix'].items():
        pct = (value / metrics['total_value']) * 100
        print(f"  {asset_class:<25} ${value:>12,.2f} ({pct:>5.1f}%)")

    print("\nAccount Breakdown:")
    print("-" * 40)
    for account, value in metrics['account_values'].items():
        pct = (value / metrics['total_value']) * 100
        print(f"  {account:<25} ${value:>12,.2f} ({pct:>5.1f}%)")

    print("\n" + "=" * 60)


def main():
    """Main entry point."""
    # Default ODS file path
    ods_file = "Portfolio.ods"

    # Allow command line argument for file path
    if len(sys.argv) > 1:
        ods_file = sys.argv[1]

    # Check if file exists
    if not Path(ods_file).exists():
        print(f"Error: File not found: {ods_file}")
        print("Usage: python portfolio_visualizer.py [path/to/portfolio.ods]")
        sys.exit(1)

    print(f"Loading portfolio data from: {ods_file}")

    # Load data
    df = load_portfolio_data(ods_file)
    print(f"Loaded {len(df)} holdings")

    # Calculate metrics
    metrics = calculate_portfolio_metrics(df)

    # Print summary
    print_summary(metrics)

    # Create visualizations
    print("\nGenerating visualizations...")
    output_file = create_visualizations(df, metrics)

    print(f"\nDashboard saved to: {output_file}")
    print("Open this file to view your portfolio visualizations!")


if __name__ == "__main__":
    main()
