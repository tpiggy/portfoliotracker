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


def load_portfolio_data(ods_file: str) -> pd.DataFrame:
    """Load investment data from an ODS file into a pandas DataFrame."""
    data = pyexcel.get_array(file_name=ods_file)

    if not data or len(data) < 2:
        raise ValueError("ODS file is empty or has no data rows")

    # First row is header
    header = data[0]
    rows = data[1:]

    df = pd.DataFrame(rows, columns=header)

    # Ensure numeric columns are proper types
    if 'Value' in df.columns:
        df['Value'] = pd.to_numeric(df['Value'], errors='coerce')
    if 'Shares' in df.columns:
        df['Shares'] = pd.to_numeric(df['Shares'], errors='coerce')
    if 'Price' in df.columns:
        df['Price'] = pd.to_numeric(df['Price'], errors='coerce')

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

    summary_text = f"{num_holdings} Holdings  |  {num_accounts} Accounts  |  {num_asset_classes} Asset Classes"
    ax1.text(5, 4, summary_text, fontsize=11, ha='center', va='center', color='#666666')

    # Top holding info
    top = metrics['top_holding']
    ax1.text(5, 2, f"Top Holding: {top['ticker']} ({top['name']})",
             fontsize=10, ha='center', va='center', color='#1565C0')
    ax1.text(5, 1, f"${top['value']:,.2f} in {top['account']}",
             fontsize=10, ha='center', va='center', color='#666666')

    # 2. Account Breakdown - Horizontal Bar Chart
    ax2 = axes[0, 1]
    account_data = metrics['account_values']
    y_pos = range(len(account_data))
    bars = ax2.barh(y_pos, account_data.values, color=[colors[i % len(colors)] for i in range(len(account_data))])
    ax2.set_yticks(y_pos)
    ax2.set_yticklabels(account_data.index)
    ax2.set_xlabel('Value ($)')
    ax2.set_title('Value by Account', fontsize=12, fontweight='bold', pad=10)

    # Add value labels on bars
    for i, (bar, val) in enumerate(zip(bars, account_data.values)):
        ax2.text(val + metrics['total_value'] * 0.01, bar.get_y() + bar.get_height()/2,
                 f'${val:,.0f}', va='center', fontsize=9)

    # Highlight largest account
    largest_idx = list(account_data.index).index(metrics['largest_account'])
    bars[largest_idx].set_color('#2E7D32')
    bars[largest_idx].set_edgecolor('#1B5E20')
    bars[largest_idx].set_linewidth(2)

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
    ods_file = "investment_portfolio.ods"

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
