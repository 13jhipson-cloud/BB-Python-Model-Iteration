#!/usr/bin/env python3
"""
Budget Comparison Script
========================
Compares forecast model outputs against the Budget consol file.

Usage:
    python compare_to_budget.py [--budget PATH] [--forecast PATH] [--output PATH]

Outputs:
    - Variance report by segment and month
    - Metrics compared: Collections, GBV, Revenue, Net Impairment
"""

import pandas as pd
import numpy as np
import argparse
import os
from datetime import datetime

# =============================================================================
# CONFIGURATION
# =============================================================================

# Segment mapping: Model segment -> Budget segment
# Adjust this if your model uses different segment names
SEGMENT_MAPPING = {
    'NON PRIME': 'Non Prime',
    'NRP-S': 'Near Prime Small',
    'NRP-M': 'Near Prime Medium',
    'NRP-L': 'Near Prime Medium',  # Combine NRP-L with Near Prime Medium
    'PRIME': 'Prime',
}

# Budget file row indices for each metric (0-indexed)
BUDGET_ROWS = {
    'Collections': {
        'Non Prime': 11,
        'Near Prime Small': 12,
        'Near Prime Medium': 13,
        'Prime': 14,
        'Total': 15,
    },
    'ClosingGBV': {
        'Non Prime': 22,
        'Near Prime Small': 23,
        'Near Prime Medium': 24,
        'Prime': 25,
        'Total': 26,
    },
    'ClosingNBV': {
        'Non Prime': 42,
        'Near Prime Small': 43,
        'Near Prime Medium': 44,
        'Prime': 45,
        'Total': 46,
    },
    'Revenue': {
        'Non Prime': 62,
        'Near Prime Small': 63,
        'Near Prime Medium': 64,
        'Prime': 65,
        'Total': 66,
    },
    'GrossImpairment': {
        'Non Prime': 73,
        'Near Prime Small': 74,
        'Near Prime Medium': 75,
        'Prime': 76,
        'Total': 77,
    },
    'DebtSaleGain': {
        'Non Prime': 105,
        'Near Prime Small': 106,
        'Near Prime Medium': 107,
        'Prime': 108,
        'Total': 109,
    },
    'NetImpairment': {
        'Non Prime': 121,
        'Near Prime Small': 122,
        'Near Prime Medium': 123,
        'Prime': 124,
        'Total': 125,
    },
}

# =============================================================================
# BUDGET LOADING
# =============================================================================

def load_budget(budget_path: str) -> pd.DataFrame:
    """
    Load budget data from Excel file and reshape to long format.

    Returns DataFrame with columns: [Date, Segment, Metric, Budget_Value]
    """
    print(f"Loading budget from: {budget_path}")

    xl = pd.ExcelFile(budget_path)
    df_raw = pd.read_excel(xl, sheet_name='P&L analysis - BB', header=None)

    # Extract dates from row 2, columns 3 onwards
    dates = pd.to_datetime(df_raw.iloc[2, 3:].values)

    # Build long-format dataframe
    records = []

    for metric, segment_rows in BUDGET_ROWS.items():
        for segment, row_idx in segment_rows.items():
            values = df_raw.iloc[row_idx, 3:].values
            for date, value in zip(dates, values):
                if pd.notna(value):
                    records.append({
                        'Date': date,
                        'Segment': segment,
                        'Metric': metric,
                        'Budget_Value': float(value)
                    })

    budget_df = pd.DataFrame(records)
    print(f"  Loaded {len(budget_df)} budget data points")
    print(f"  Date range: {budget_df['Date'].min()} to {budget_df['Date'].max()}")

    return budget_df


# =============================================================================
# FORECAST LOADING
# =============================================================================

def load_forecast(forecast_path: str) -> pd.DataFrame:
    """
    Load forecast data from model output and reshape to match budget format.

    Returns DataFrame with columns: [Date, Segment, Metric, Forecast_Value]
    """
    print(f"Loading forecast from: {forecast_path}")

    # Check if it's a directory or file
    if os.path.isdir(forecast_path):
        summary_path = os.path.join(forecast_path, 'Forecast_Summary.xlsx')
        impairment_path = os.path.join(forecast_path, 'Impairment_Analysis.xlsx')
    else:
        summary_path = forecast_path
        impairment_path = forecast_path.replace('Forecast_Summary', 'Impairment_Analysis')

    # Load forecast summary
    df_summary = pd.read_excel(summary_path, sheet_name='Summary')
    df_summary['ForecastMonth'] = pd.to_datetime(df_summary['ForecastMonth'])

    # Load impairment details for gross impairment and debt sale metrics
    df_impairment = pd.read_excel(impairment_path, sheet_name='Impairment_Detail')
    df_impairment['ForecastMonth'] = pd.to_datetime(df_impairment['ForecastMonth'])

    # Map model segments to budget segments
    df_summary['BudgetSegment'] = df_summary['Segment'].map(SEGMENT_MAPPING)
    df_impairment['BudgetSegment'] = df_impairment['Segment'].map(SEGMENT_MAPPING)

    # Aggregate summary by budget segment and month
    agg_df = df_summary.groupby(['ForecastMonth', 'BudgetSegment']).agg({
        'OpeningGBV': 'sum',
        'ClosingGBV': 'sum',
        'ClosingNBV': 'sum',
        'InterestRevenue': 'sum',
        'Coll_Principal': 'sum',
        'Coll_Interest': 'sum',
        'WO_DebtSold': 'sum',
        'WO_Other': 'sum',
        'Net_Impairment': 'sum',
    }).reset_index()

    # Aggregate impairment by budget segment and month
    imp_agg = df_impairment.groupby(['ForecastMonth', 'BudgetSegment']).agg({
        'Gross_Impairment_ExcludingDS': 'sum',
        'Debt_Sale_Impact': 'sum',
        'Net_Impairment': 'sum',
    }).reset_index()

    # Merge impairment data
    agg_df = agg_df.merge(
        imp_agg[['ForecastMonth', 'BudgetSegment', 'Gross_Impairment_ExcludingDS', 'Debt_Sale_Impact']],
        on=['ForecastMonth', 'BudgetSegment'],
        how='left'
    )

    # Calculate derived metrics
    # Collections in model are negative (reduce GBV), budget shows positive (cash received)
    agg_df['Collections'] = -(agg_df['Coll_Principal'] + agg_df['Coll_Interest'])
    agg_df['Revenue'] = agg_df['InterestRevenue']
    # Gross impairment (negative = charge, as per reporting convention)
    agg_df['GrossImpairment'] = agg_df['Gross_Impairment_ExcludingDS']
    # Debt sale gain (positive = benefit)
    # With new sign convention, Debt_Sale_Impact already represents gain directly
    agg_df['DebtSaleGain'] = agg_df['Debt_Sale_Impact']
    # Net impairment
    agg_df['NetImpairment'] = agg_df['Net_Impairment']

    # Reshape to long format
    records = []
    metrics_to_extract = ['Collections', 'ClosingGBV', 'ClosingNBV', 'Revenue',
                          'GrossImpairment', 'DebtSaleGain', 'NetImpairment']

    for _, row in agg_df.iterrows():
        for metric in metrics_to_extract:
            records.append({
                'Date': row['ForecastMonth'],
                'Segment': row['BudgetSegment'],
                'Metric': metric,
                'Forecast_Value': row[metric]
            })

    # Also calculate totals
    total_df = agg_df.groupby('ForecastMonth').agg({
        'Collections': 'sum',
        'ClosingGBV': 'sum',
        'ClosingNBV': 'sum',
        'Revenue': 'sum',
        'GrossImpairment': 'sum',
        'DebtSaleGain': 'sum',
        'NetImpairment': 'sum',
    }).reset_index()

    for _, row in total_df.iterrows():
        for metric in metrics_to_extract:
            records.append({
                'Date': row['ForecastMonth'],
                'Segment': 'Total',
                'Metric': metric,
                'Forecast_Value': row[metric]
            })

    forecast_df = pd.DataFrame(records)
    print(f"  Loaded {len(forecast_df)} forecast data points")
    print(f"  Date range: {forecast_df['Date'].min()} to {forecast_df['Date'].max()}")

    return forecast_df


# =============================================================================
# COMPARISON
# =============================================================================

def compare_budget_vs_forecast(budget_df: pd.DataFrame, forecast_df: pd.DataFrame,
                               match_dates: bool = True) -> pd.DataFrame:
    """
    Compare budget vs forecast and calculate variances.

    Args:
        budget_df: Budget data in long format
        forecast_df: Forecast data in long format
        match_dates: If True, only compare dates that exist in both datasets

    Returns DataFrame with variance analysis.
    """
    print("Comparing budget vs forecast...")

    if match_dates:
        # Find common dates
        budget_dates = set(budget_df['Date'].unique())
        forecast_dates = set(forecast_df['Date'].unique())
        common_dates = budget_dates & forecast_dates
        print(f"  Common date range: {min(common_dates)} to {max(common_dates)} ({len(common_dates)} months)")

        budget_df = budget_df[budget_df['Date'].isin(common_dates)]
        forecast_df = forecast_df[forecast_df['Date'].isin(common_dates)]

    # Merge on Date, Segment, Metric
    merged = pd.merge(
        budget_df,
        forecast_df,
        on=['Date', 'Segment', 'Metric'],
        how='outer'
    )

    # Calculate variance metrics
    merged['Variance'] = merged['Forecast_Value'] - merged['Budget_Value']
    merged['Variance_Pct'] = np.where(
        merged['Budget_Value'] != 0,
        (merged['Forecast_Value'] / merged['Budget_Value'] - 1) * 100,
        np.nan
    )

    # Flag large variances (>5%)
    merged['Large_Variance'] = np.abs(merged['Variance_Pct']) > 5

    # Sort for readability
    merged = merged.sort_values(['Metric', 'Segment', 'Date'])

    print(f"  Total comparison rows: {len(merged)}")
    print(f"  Rows with large variance (>5%): {merged['Large_Variance'].sum()}")

    return merged


def generate_summary(comparison_df: pd.DataFrame) -> pd.DataFrame:
    """
    Generate summary statistics by metric and segment.
    """
    summary = comparison_df.groupby(['Metric', 'Segment']).agg({
        'Budget_Value': 'sum',
        'Forecast_Value': 'sum',
        'Variance': 'sum',
        'Large_Variance': 'sum'
    }).reset_index()

    summary['Total_Variance_Pct'] = np.where(
        summary['Budget_Value'] != 0,
        (summary['Forecast_Value'] / summary['Budget_Value'] - 1) * 100,
        np.nan
    )

    summary = summary.rename(columns={'Large_Variance': 'Months_With_Large_Variance'})

    return summary


# =============================================================================
# OUTPUT
# =============================================================================

def save_comparison_report(comparison_df: pd.DataFrame, summary_df: pd.DataFrame, output_path: str):
    """
    Save comparison results to Excel.
    """
    print(f"Saving comparison report to: {output_path}")

    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        # Summary sheet
        summary_df.to_excel(writer, sheet_name='Summary', index=False)

        # Detailed comparison by metric
        for metric in comparison_df['Metric'].unique():
            metric_df = comparison_df[comparison_df['Metric'] == metric].copy()
            # Pivot for readability
            pivot = metric_df.pivot_table(
                index='Date',
                columns='Segment',
                values=['Budget_Value', 'Forecast_Value', 'Variance_Pct'],
                aggfunc='first'
            )
            # Flatten column names
            pivot.columns = [f"{col[1]}_{col[0]}" for col in pivot.columns]
            pivot = pivot.reset_index()

            sheet_name = metric[:31]  # Excel sheet name limit
            pivot.to_excel(writer, sheet_name=sheet_name, index=False)

        # Full detail
        comparison_df.to_excel(writer, sheet_name='Full_Detail', index=False)

    print(f"  Report saved successfully")


def print_summary(summary_df: pd.DataFrame):
    """
    Print summary to console.
    """
    print("\n" + "=" * 80)
    print("BUDGET VS FORECAST COMPARISON SUMMARY")
    print("=" * 80)

    # Print by metric
    for metric in summary_df['Metric'].unique():
        print(f"\n{metric}:")
        metric_summary = summary_df[summary_df['Metric'] == metric]

        for _, row in metric_summary.iterrows():
            segment = row['Segment']
            budget = row['Budget_Value'] / 1e6  # Convert to millions
            forecast = row['Forecast_Value'] / 1e6
            var_pct = row['Total_Variance_Pct']
            flag = "⚠️" if abs(var_pct) > 5 else "✓"

            print(f"  {segment:20} Budget: £{budget:12,.2f}m  Forecast: £{forecast:12,.2f}m  Var: {var_pct:+7.2f}% {flag}")


# =============================================================================
# MAIN
# =============================================================================

def main():
    parser = argparse.ArgumentParser(description='Compare forecast vs budget')
    parser.add_argument('--budget', default='Budget consol file.xlsx',
                        help='Path to budget Excel file')
    parser.add_argument('--forecast', default='output',
                        help='Path to forecast output directory or Forecast_Summary.xlsx')
    parser.add_argument('--output', default='output/Budget_Comparison.xlsx',
                        help='Path for output comparison report')

    args = parser.parse_args()

    # Load data
    budget_df = load_budget(args.budget)
    forecast_df = load_forecast(args.forecast)

    # Compare
    comparison_df = compare_budget_vs_forecast(budget_df, forecast_df)
    summary_df = generate_summary(comparison_df)

    # Output
    save_comparison_report(comparison_df, summary_df, args.output)
    print_summary(summary_df)

    print("\n" + "=" * 80)
    print("Comparison complete!")
    print("=" * 80)

    return comparison_df, summary_df


if __name__ == '__main__':
    main()
