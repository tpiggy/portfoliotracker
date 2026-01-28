#!/usr/bin/env python3
"""
Investment Portfolio Visualizer

Reads an investment tracking ODS file and creates visualizations including:
- Total Portfolio Value
- Largest Account
- Top Holding
- Asset Mix
- Risk Analysis & Concentration Warnings
- Consolidated Holdings View
- Daily Performance Tracking
- Rebalancing Suggestions
"""

import sys
import pyexcel
import pandas as pd
import matplotlib.pyplot as plt
from pathlib import Path
from datetime import datetime


# Target allocation for rebalancing (customize as needed)
TARGET_ALLOCATION = {
    'Equities': 0.30,
    'Mining - Copper': 0.10,
    'Mining - Coal': 0.10,
    'Mining - Gold': 0.05,
    'Mining - Services': 0.05,
    'Energy - Oil & Gas': 0.05,
    'Energy - Pipelines': 0.05,
    'Real Estate': 0.05,
    'Cryptocurrency': 0.10,
    'Cash': 0.10,
    'Other': 0.05,
}

# Risk thresholds
MAX_SINGLE_HOLDING_PCT = 0.15  # 15% max for any single holding
MAX_SECTOR_PCT = 0.30  # 30% max for any single sector
MAX_CRYPTO_PCT = 0.15  # 15% max for crypto
MIN_CASH_PCT = 0.05  # 5% minimum cash


# Asset class mappings based on ticker/name patterns
ASSET_CLASS_RULES = {
    'CASH$': 'Cash',
    # Mining
    'MMA': 'Mining - Copper',
    'AMR': 'Mining - Coal',
    'CVE': 'Mining - Coal',
    'WM': 'Mining - Gold',
    'FAR': 'Mining - Services',
    # Energy
    'TOU': 'Energy - Oil & Gas',
    'PPL': 'Energy - Pipelines',
    'ET': 'Energy - Pipelines',
    # Real Estate
    'LAND': 'Real Estate',
    # Consumer
    'LULU': 'Consumer Discretionary',
    'GURU': 'Consumer Staples',
    'PZZA': 'Consumer Discretionary',
    'PZA': 'Consumer Discretionary',
    'ABEV': 'Consumer Staples',
    # Tech / Aerospace
    'RKLB': 'Aerospace & Defense',
    # Healthcare / Biotech
    'CRSP': 'Biotech',
    'BEAM': 'Biotech',
    # Transportation
    'CNR': 'Transportation',
    'CNI': 'Transportation',
    # Financials
    'BNS': 'Financials',
    'TD': 'Financials',
    'RY': 'Financials',
    # Crypto
    'BTC': 'Cryptocurrency',
    'CARDANO': 'Cryptocurrency',
    'ADA': 'Cryptocurrency',
    'ETH': 'Cryptocurrency',
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


def parse_sheet_data(data: list, account_name: str) -> list:
    """Parse data from a single sheet and return list of row dicts."""
    rows = []
    header = ['Name', 'Ticker', 'Price', 'Shares', 'Value', 'Percentage', 'Change', 'Extra1', 'Extra2']

    # Find header row if present (to know where data ends)
    header_row_idx = len(data)
    for i, row in enumerate(data):
        if len(row) > 0 and row[0] == 'Name' and len(row) > 1 and row[1] == 'Ticker':
            header_row_idx = i
            break

    # Parse data rows (before header row if present)
    for row in data[:header_row_idx]:
        # Skip empty rows or rows without valid data
        if len(row) < 5 or not row[0] or not row[1]:
            continue

        # Skip metadata rows (like "USD conv rate", totals, etc.)
        ticker = str(row[1]).strip().upper()
        name = str(row[0]).strip().lower()

        # Skip internal references and totals
        if ticker in ['', 'TFSA', 'RRSP', 'NAME', 'TICKER']:
            continue
        if name in ['tfsa', 'rrsp', 'total', '']:
            continue

        # Check if this is a crypto row with different format
        # Crypto format: [Ticker, Price, Shares, Label, Value, ...]
        col3_str = str(row[3]).strip().lower() if len(row) > 3 else ''
        if 'total' in col3_str:
            # This is crypto format: Name is ticker, col[1] is price, col[2] is shares, col[4] is value
            try:
                value = float(row[4]) if row[4] else 0
                if value > 0:
                    row_dict = {
                        'Name': str(row[0]),
                        'Ticker': str(row[0]).upper(),
                        'Price': row[1],
                        'Shares': row[2],
                        'Value': value,
                        'Percentage': '',
                        'Change': '',
                        'Extra1': '',
                        'Extra2': '',
                        'Account': account_name
                    }
                    rows.append(row_dict)
            except (ValueError, TypeError):
                pass
            continue

        try:
            value = float(row[4]) if row[4] else 0
            if value > 0:
                row_data = row[:9] if len(row) >= 9 else row + [''] * (9 - len(row))
                row_dict = dict(zip(header, row_data))
                row_dict['Account'] = account_name
                rows.append(row_dict)
        except (ValueError, TypeError):
            continue

    return rows


def load_portfolio_data(ods_file: str) -> pd.DataFrame:
    """Load investment data from an ODS file into a pandas DataFrame.

    Supports multi-sheet ODS files where each sheet represents a different account.
    """
    # Get all sheet names
    book = pyexcel.get_book(file_name=ods_file)
    sheet_names = book.sheet_names()

    all_rows = []

    if len(sheet_names) > 1:
        # Multi-sheet format: each sheet is an account
        print(f"Found {len(sheet_names)} sheets: {', '.join(sheet_names)}")
        for sheet_name in sheet_names:
            data = pyexcel.get_array(file_name=ods_file, sheet_name=sheet_name)
            if data:
                rows = parse_sheet_data(data, sheet_name)
                all_rows.extend(rows)
                print(f"  - {sheet_name}: {len(rows)} holdings")
    else:
        # Single sheet format
        data = pyexcel.get_array(file_name=ods_file)
        if data:
            all_rows = parse_sheet_data(data, 'Main Portfolio')

    if not all_rows:
        raise ValueError("ODS file is empty or has no valid data rows")

    df = pd.DataFrame(all_rows)

    # Ensure numeric columns are proper types
    for col in ['Value', 'Shares', 'Price', 'Percentage', 'Change']:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

    # Add Asset Class column if not present
    if 'Asset Class' not in df.columns:
        df['Asset Class'] = df.apply(lambda row: infer_asset_class(row.get('Ticker', ''), row.get('Name', '')), axis=1)

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


def analyze_risk(df: pd.DataFrame, metrics: dict) -> dict:
    """Analyze portfolio risk and generate warnings."""
    risk = {
        'warnings': [],
        'concentration_score': 0,
        'diversification_rating': '',
        'hhi': 0,  # Herfindahl-Hirschman Index
    }

    total_value = metrics['total_value']

    # Calculate HHI (sum of squared market shares) - lower is more diversified
    holdings_pct = (df['Value'] / total_value * 100) ** 2
    risk['hhi'] = holdings_pct.sum()

    # Diversification rating based on HHI
    if risk['hhi'] < 1000:
        risk['diversification_rating'] = 'Excellent'
        risk['concentration_score'] = 10
    elif risk['hhi'] < 1500:
        risk['diversification_rating'] = 'Good'
        risk['concentration_score'] = 8
    elif risk['hhi'] < 2500:
        risk['diversification_rating'] = 'Moderate'
        risk['concentration_score'] = 6
    elif risk['hhi'] < 4000:
        risk['diversification_rating'] = 'Concentrated'
        risk['concentration_score'] = 4
    else:
        risk['diversification_rating'] = 'Highly Concentrated'
        risk['concentration_score'] = 2

    # Check single holding concentration
    for _, row in df.iterrows():
        pct = row['Value'] / total_value
        if pct > MAX_SINGLE_HOLDING_PCT:
            risk['warnings'].append({
                'type': 'HIGH_CONCENTRATION',
                'severity': 'high' if pct > 0.25 else 'medium',
                'message': f"{row['Ticker']} is {pct*100:.1f}% of portfolio (>{MAX_SINGLE_HOLDING_PCT*100:.0f}% threshold)"
            })

    # Check sector concentration
    for sector, value in metrics['asset_mix'].items():
        pct = value / total_value
        if pct > MAX_SECTOR_PCT:
            risk['warnings'].append({
                'type': 'SECTOR_CONCENTRATION',
                'severity': 'high' if pct > 0.40 else 'medium',
                'message': f"{sector} is {pct*100:.1f}% of portfolio (>{MAX_SECTOR_PCT*100:.0f}% threshold)"
            })

    # Check crypto exposure
    crypto_value = metrics['asset_mix'].get('Cryptocurrency', 0)
    crypto_pct = crypto_value / total_value
    if crypto_pct > MAX_CRYPTO_PCT:
        risk['warnings'].append({
            'type': 'CRYPTO_EXPOSURE',
            'severity': 'high' if crypto_pct > 0.30 else 'medium',
            'message': f"Cryptocurrency is {crypto_pct*100:.1f}% of portfolio (>{MAX_CRYPTO_PCT*100:.0f}% recommended max)"
        })

    # Check cash levels
    cash_value = metrics['asset_mix'].get('Cash', 0)
    cash_pct = cash_value / total_value
    if cash_pct < MIN_CASH_PCT:
        risk['warnings'].append({
            'type': 'LOW_CASH',
            'severity': 'low',
            'message': f"Cash is only {cash_pct*100:.1f}% of portfolio (<{MIN_CASH_PCT*100:.0f}% recommended min)"
        })

    return risk


def get_consolidated_holdings(df: pd.DataFrame) -> pd.DataFrame:
    """Consolidate same ticker across all accounts."""
    consolidated = df.groupby('Ticker').agg({
        'Name': 'first',
        'Value': 'sum',
        'Shares': 'sum',
        'Price': 'first',
        'Asset Class': 'first',
        'Account': lambda x: ', '.join(sorted(set(x)))
    }).reset_index()

    consolidated = consolidated.sort_values('Value', ascending=False)
    consolidated['Pct'] = consolidated['Value'] / consolidated['Value'].sum() * 100

    return consolidated


def calculate_rebalancing(metrics: dict) -> pd.DataFrame:
    """Calculate rebalancing suggestions based on target allocation."""
    total_value = metrics['total_value']
    current_allocation = metrics['asset_mix']

    rebalance_data = []
    for sector, target_pct in TARGET_ALLOCATION.items():
        current_value = current_allocation.get(sector, 0)
        current_pct = current_value / total_value

        target_value = total_value * target_pct
        difference = target_value - current_value
        diff_pct = target_pct - current_pct

        if abs(diff_pct) > 0.01:  # Only show if difference > 1%
            rebalance_data.append({
                'Sector': sector,
                'Current %': current_pct * 100,
                'Target %': target_pct * 100,
                'Diff %': diff_pct * 100,
                'Action': 'BUY' if difference > 0 else 'SELL',
                'Amount': abs(difference)
            })

    # Add sectors not in target but in portfolio
    for sector, value in current_allocation.items():
        if sector not in TARGET_ALLOCATION:
            current_pct = value / total_value
            if current_pct > 0.01:
                rebalance_data.append({
                    'Sector': sector,
                    'Current %': current_pct * 100,
                    'Target %': 0,
                    'Diff %': -current_pct * 100,
                    'Action': 'REVIEW',
                    'Amount': value
                })

    return pd.DataFrame(rebalance_data).sort_values('Diff %', key=abs, ascending=False)


def calculate_daily_performance(df: pd.DataFrame) -> dict:
    """Calculate daily performance metrics from Change column."""
    perf = {
        'daily_change_total': 0,
        'daily_change_pct': 0,
        'gainers': [],
        'losers': [],
    }

    if 'Change' not in df.columns:
        return perf

    # Filter rows with valid change data
    df_with_change = df[df['Change'] != 0].copy()

    if df_with_change.empty:
        return perf

    # Calculate daily dollar change for each holding
    # Change is typically % change, so: dollar_change = value * change / (100 + change)
    # Or approximate: dollar_change = value * (change/100) for small changes
    df_with_change['Dollar_Change'] = df_with_change.apply(
        lambda row: row['Value'] * (row['Change'] / 100) / (1 + row['Change'] / 100)
        if row['Change'] != 0 else 0, axis=1
    )

    perf['daily_change_total'] = df_with_change['Dollar_Change'].sum()

    # Previous day's value
    prev_value = df['Value'].sum() - perf['daily_change_total']
    if prev_value > 0:
        perf['daily_change_pct'] = (perf['daily_change_total'] / prev_value) * 100

    # Top gainers and losers
    sorted_by_change = df_with_change.sort_values('Change', ascending=False)

    gainers = sorted_by_change[sorted_by_change['Change'] > 0].head(5)
    perf['gainers'] = [
        {'ticker': row['Ticker'], 'change': row['Change'], 'dollar': row['Dollar_Change']}
        for _, row in gainers.iterrows()
    ]

    losers = sorted_by_change[sorted_by_change['Change'] < 0].tail(5)
    perf['losers'] = [
        {'ticker': row['Ticker'], 'change': row['Change'], 'dollar': row['Dollar_Change']}
        for _, row in losers.iterrows()
    ]

    return perf


def create_visualizations(df: pd.DataFrame, metrics: dict, risk: dict, perf: dict, output_dir: str = "."):
    """Create and save portfolio visualizations."""
    output_path = Path(output_dir)
    output_path.mkdir(exist_ok=True)

    # Set up the style
    plt.style.use('seaborn-v0_8-whitegrid')
    colors = plt.cm.Set3.colors

    # Create a figure with 6 subplots (2x3)
    fig, axes = plt.subplots(2, 3, figsize=(18, 12))
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

    # 3. Risk Analysis Panel
    ax3 = axes[0, 2]
    ax3.set_xlim(0, 10)
    ax3.set_ylim(0, 10)
    ax3.axis('off')

    ax3.text(5, 9.5, 'Risk Analysis', fontsize=12, fontweight='bold', ha='center', va='center')

    # Diversification rating
    rating_color = {'Excellent': '#2E7D32', 'Good': '#4CAF50', 'Moderate': '#FFC107',
                    'Concentrated': '#FF9800', 'Highly Concentrated': '#F44336'}
    color = rating_color.get(risk['diversification_rating'], '#666666')
    ax3.text(5, 8, f"Diversification: {risk['diversification_rating']}", fontsize=11,
             ha='center', va='center', color=color, fontweight='bold')
    ax3.text(5, 7, f"HHI Score: {risk['hhi']:.0f}", fontsize=9, ha='center', va='center', color='#666666')

    # Warnings
    y_pos = 5.5
    if risk['warnings']:
        ax3.text(5, 6, 'Warnings:', fontsize=10, ha='center', va='center', fontweight='bold', color='#D32F2F')
        for warning in risk['warnings'][:4]:  # Show max 4 warnings
            severity_color = '#D32F2F' if warning['severity'] == 'high' else '#FF9800' if warning['severity'] == 'medium' else '#666666'
            # Truncate long messages
            msg = warning['message'][:40] + '...' if len(warning['message']) > 40 else warning['message']
            ax3.text(5, y_pos, f"• {msg}", fontsize=8, ha='center', va='center', color=severity_color)
            y_pos -= 0.9
    else:
        ax3.text(5, 5, 'No warnings', fontsize=10, ha='center', va='center', color='#2E7D32')

    # 4. Asset Mix - Pie Chart
    ax4 = axes[1, 0]
    asset_data = metrics['asset_mix']

    # Create pie chart with percentages
    wedges, texts, autotexts = ax4.pie(
        asset_data.values,
        labels=asset_data.index,
        autopct=lambda pct: f'{pct:.1f}%' if pct > 3 else '',
        colors=[colors[i % len(colors)] for i in range(len(asset_data))],
        explode=[0.02] * len(asset_data),
        startangle=90
    )
    ax4.set_title('Asset Mix', fontsize=12, fontweight='bold', pad=10)

    # Style the percentage labels
    for autotext in autotexts:
        autotext.set_fontsize(9)
        autotext.set_fontweight('bold')

    # 5. Top 10 Holdings - Horizontal Bar Chart
    ax5 = axes[1, 1]
    top_holdings = metrics['top_holdings']
    y_pos = range(len(top_holdings))

    # Create labels with ticker and name
    labels = [f"{row['Ticker']}" for _, row in top_holdings.iterrows()]

    bars = ax5.barh(y_pos, top_holdings['Value'].values,
                    color=[colors[i % len(colors)] for i in range(len(top_holdings))])
    ax5.set_yticks(y_pos)
    ax5.set_yticklabels(labels)
    ax5.set_xlabel('Value ($)')
    ax5.set_title('Top 10 Holdings', fontsize=12, fontweight='bold', pad=10)

    # Add value labels
    for bar, val in zip(bars, top_holdings['Value'].values):
        ax5.text(val + metrics['total_value'] * 0.005, bar.get_y() + bar.get_height()/2,
                 f'${val:,.0f}', va='center', fontsize=9)

    # Highlight top holding
    bars[0].set_color('#1565C0')
    bars[0].set_edgecolor('#0D47A1')
    bars[0].set_linewidth(2)

    ax5.invert_yaxis()

    # 6. Daily Performance Panel
    ax6 = axes[1, 2]
    ax6.set_xlim(0, 10)
    ax6.set_ylim(0, 10)
    ax6.axis('off')

    ax6.text(5, 9.5, 'Daily Performance', fontsize=12, fontweight='bold', ha='center', va='center')

    # Daily change
    change_color = '#2E7D32' if perf['daily_change_total'] >= 0 else '#D32F2F'
    change_sign = '+' if perf['daily_change_total'] >= 0 else ''
    ax6.text(5, 8, f"{change_sign}${perf['daily_change_total']:,.2f}", fontsize=18, fontweight='bold',
             ha='center', va='center', color=change_color)
    ax6.text(5, 7, f"({change_sign}{perf['daily_change_pct']:.2f}%)", fontsize=11,
             ha='center', va='center', color=change_color)

    # Top gainers
    y_pos = 5.5
    if perf['gainers']:
        ax6.text(2.5, 6, 'Top Gainers', fontsize=9, ha='center', va='center', fontweight='bold', color='#2E7D32')
        for g in perf['gainers'][:3]:
            ax6.text(2.5, y_pos, f"{g['ticker']}: +{g['change']:.1f}%", fontsize=8,
                     ha='center', va='center', color='#2E7D32')
            y_pos -= 0.7

    # Top losers
    y_pos = 5.5
    if perf['losers']:
        ax6.text(7.5, 6, 'Top Losers', fontsize=9, ha='center', va='center', fontweight='bold', color='#D32F2F')
        for l in perf['losers'][:3]:
            ax6.text(7.5, y_pos, f"{l['ticker']}: {l['change']:.1f}%", fontsize=8,
                     ha='center', va='center', color='#D32F2F')
            y_pos -= 0.7

    plt.tight_layout()

    # Save the combined dashboard
    dashboard_path = output_path / 'portfolio_dashboard.png'
    plt.savefig(dashboard_path, dpi=150, bbox_inches='tight', facecolor='white')
    print(f"Saved: {dashboard_path}")

    plt.close()

    return str(dashboard_path)


def print_summary(metrics: dict, risk: dict, perf: dict, consolidated: pd.DataFrame, rebalance: pd.DataFrame):
    """Print a comprehensive text summary of the portfolio."""
    print("\n" + "=" * 70)
    print("INVESTMENT PORTFOLIO SUMMARY")
    print(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}")
    print("=" * 70)

    print(f"\n{'Total Portfolio Value:':<25} ${metrics['total_value']:>15,.2f}")

    # Daily performance
    change_sign = '+' if perf['daily_change_total'] >= 0 else ''
    print(f"{'Daily Change:':<25} {change_sign}${perf['daily_change_total']:>14,.2f} ({change_sign}{perf['daily_change_pct']:.2f}%)")

    print(f"\n{'Largest Account:':<25} {metrics['largest_account']}")
    print(f"{'   Value:':<25} ${metrics['largest_account_value']:>15,.2f}")

    top = metrics['top_holding']
    print(f"\n{'Top Holding:':<25} {top['ticker']} - {top['name']}")
    print(f"{'   Value:':<25} ${top['value']:>15,.2f}")
    print(f"{'   Account:':<25} {top['account']}")

    # Risk Analysis
    print("\n" + "-" * 70)
    print("RISK ANALYSIS")
    print("-" * 70)
    print(f"{'Diversification Rating:':<25} {risk['diversification_rating']}")
    print(f"{'HHI Score:':<25} {risk['hhi']:.0f} (lower is more diversified)")

    if risk['warnings']:
        print(f"\n{'Warnings:':<25} {len(risk['warnings'])} issues found")
        for w in risk['warnings']:
            severity_icon = '!!' if w['severity'] == 'high' else '!' if w['severity'] == 'medium' else '~'
            print(f"  [{severity_icon}] {w['message']}")
    else:
        print(f"\n{'Warnings:':<25} None - portfolio looks healthy!")

    # Consolidated Holdings (top 10)
    print("\n" + "-" * 70)
    print("CONSOLIDATED HOLDINGS (Same ticker across accounts)")
    print("-" * 70)
    print(f"{'Ticker':<10} {'Name':<30} {'Value':>15} {'Pct':>8} {'Accounts'}")
    print("-" * 70)
    for _, row in consolidated.head(10).iterrows():
        name = str(row['Name'])[:28] + '..' if len(str(row['Name'])) > 30 else str(row['Name'])
        accounts = str(row['Account'])[:20] + '..' if len(str(row['Account'])) > 22 else str(row['Account'])
        print(f"{row['Ticker']:<10} {name:<30} ${row['Value']:>14,.2f} {row['Pct']:>6.1f}%  {accounts}")

    # Asset Mix
    print("\n" + "-" * 70)
    print("ASSET MIX")
    print("-" * 70)
    for asset_class, value in metrics['asset_mix'].items():
        pct = (value / metrics['total_value']) * 100
        print(f"  {asset_class:<25} ${value:>12,.2f} ({pct:>5.1f}%)")

    # Account Breakdown
    print("\n" + "-" * 70)
    print("ACCOUNT BREAKDOWN")
    print("-" * 70)
    for account, value in metrics['account_values'].items():
        pct = (value / metrics['total_value']) * 100
        print(f"  {account:<25} ${value:>12,.2f} ({pct:>5.1f}%)")

    # Rebalancing Suggestions
    if not rebalance.empty:
        print("\n" + "-" * 70)
        print("REBALANCING SUGGESTIONS")
        print("-" * 70)
        print(f"{'Sector':<25} {'Current':>8} {'Target':>8} {'Action':>8} {'Amount':>12}")
        print("-" * 70)
        for _, row in rebalance.head(8).iterrows():
            print(f"{row['Sector']:<25} {row['Current %']:>7.1f}% {row['Target %']:>7.1f}% {row['Action']:>8} ${row['Amount']:>11,.0f}")

    print("\n" + "=" * 70)


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

    # Calculate core metrics
    metrics = calculate_portfolio_metrics(df)

    # Run risk analysis
    risk = analyze_risk(df, metrics)

    # Calculate daily performance
    perf = calculate_daily_performance(df)

    # Get consolidated holdings view
    consolidated = get_consolidated_holdings(df)

    # Calculate rebalancing suggestions
    rebalance = calculate_rebalancing(metrics)

    # Print comprehensive summary
    print_summary(metrics, risk, perf, consolidated, rebalance)

    # Create visualizations
    print("\nGenerating visualizations...")
    output_file = create_visualizations(df, metrics, risk, perf)

    print(f"\nDashboard saved to: {output_file}")
    print("Open this file to view your portfolio visualizations!")


if __name__ == "__main__":
    main()
