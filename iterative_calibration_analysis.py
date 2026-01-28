#!/usr/bin/env python3
"""
Iterative Calibration Analysis

This script runs the backbook forecast model with different methodology configurations,
compares results to budget, and generates detailed analysis to inform modeling decisions.

Each iteration is saved with documentation of:
- The methodology used
- The variance analysis vs budget
- The rationale for any adjustments
"""

import pandas as pd
import numpy as np
import subprocess
import os
import sys
from datetime import datetime

# Paths
BASE_DIR = '/home/user/BB-Python-Model'
FACT_RAW = os.path.join(BASE_DIR, 'sample_data', 'Fact_Raw_Full.csv')
BUDGET_FILE = os.path.join(BASE_DIR, 'sample_data', 'Budget consol file.xlsx')
OUTPUT_DIR = os.path.join(BASE_DIR, 'calibration_iterations')

# Segment name mapping (budget uses different names)
SEGMENT_MAP = {
    'Non Prime': 'NON PRIME',
    'Near Prime Small': 'NRP-S',
    'Near Prime Medium': 'NRP-M',
    'Near Prime Large': 'NRP-L',
    'Prime': 'PRIME'
}
SEGMENT_MAP_REVERSE = {v: k for k, v in SEGMENT_MAP.items()}


def load_budget_data():
    """Load and process budget data for comparison."""
    # Read the P&L analysis sheet
    budget_df = pd.read_excel(BUDGET_FILE, sheet_name='P&L analysis')

    # The budget file has segments in columns starting from column B
    # First column is the metric names
    # Structure: Row contains metric name, then values for each segment/month combination

    # Let's read it more carefully
    budget_raw = pd.read_excel(BUDGET_FILE, sheet_name='P&L analysis', header=None)

    # Find the header row (contains segment names)
    # Typically row 0 or 1
    # Look for 'Non Prime' or similar
    header_row = None
    for i in range(min(10, len(budget_raw))):
        row_values = budget_raw.iloc[i].astype(str).tolist()
        if any('Non Prime' in str(v) or 'Near Prime' in str(v) for v in row_values):
            header_row = i
            break

    if header_row is None:
        print("Warning: Could not find header row in budget file")
        return None

    # Read with proper header
    budget_df = pd.read_excel(BUDGET_FILE, sheet_name='P&L analysis', header=header_row)

    return budget_df


def load_budget_metrics():
    """Extract budget metrics by segment and month."""
    budget_raw = pd.read_excel(BUDGET_FILE, sheet_name='P&L analysis', header=None)

    # Based on the file structure, parse the relevant sections
    # The file has metrics in rows and segment/month combinations in columns

    budget_data = []

    # Read the sheet with headers
    try:
        # Try to parse the structured data
        df = pd.read_excel(BUDGET_FILE, sheet_name='P&L analysis')

        # Find columns that look like dates or segments
        # This file typically has a multi-level structure

        # For now, return what we can parse
        return df
    except Exception as e:
        print(f"Error loading budget: {e}")
        return None


def run_forecast(methodology_file, output_prefix):
    """Run the backbook forecast model with given methodology."""
    cmd = [
        'python', os.path.join(BASE_DIR, 'backbook_forecast.py'),
        '--fact-raw', FACT_RAW,
        '--methodology', methodology_file,
        '--output-prefix', output_prefix,
        '--months', '12'
    ]

    print(f"Running forecast with methodology: {methodology_file}")
    result = subprocess.run(cmd, capture_output=True, text=True, cwd=BASE_DIR)

    if result.returncode != 0:
        print(f"Error running forecast: {result.stderr}")
        return None

    print(result.stdout[-2000:] if len(result.stdout) > 2000 else result.stdout)
    return True


def load_forecast_results(output_prefix):
    """Load the forecast results from CSV files."""
    # The model outputs several files
    forecast_file = f"{output_prefix}_forecast.csv"

    if os.path.exists(forecast_file):
        return pd.read_csv(forecast_file)

    # Try in output directory
    forecast_file = os.path.join(BASE_DIR, 'output', f"{output_prefix}_forecast.csv")
    if os.path.exists(forecast_file):
        return pd.read_csv(forecast_file)

    # List what files exist
    print(f"Looking for forecast output files...")
    for f in os.listdir(os.path.join(BASE_DIR, 'output')):
        if output_prefix in f or 'forecast' in f.lower():
            print(f"  Found: {f}")

    return None


def analyze_historical_rates(fact_raw_path):
    """Analyze historical rate patterns from actuals to inform methodology."""
    df = pd.read_csv(fact_raw_path)

    # Parse dates
    df['Date'] = pd.to_datetime(df['Date'])

    # Get the latest 12 months of data
    max_date = df['Date'].max()
    min_date = max_date - pd.DateOffset(months=12)
    recent_df = df[df['Date'] >= min_date].copy()

    analysis = {}

    # Calculate rate metrics by segment
    for segment in df['Segment'].unique():
        seg_df = recent_df[recent_df['Segment'] == segment].copy()

        if len(seg_df) == 0:
            continue

        # Calculate average rates
        seg_analysis = {}

        # Collection rate = Collections / OpeningGBV
        if 'Coll_Principal' in seg_df.columns and 'OpeningGBV' in seg_df.columns:
            seg_df['coll_rate'] = seg_df['Coll_Principal'].abs() / seg_df['OpeningGBV'].replace(0, np.nan)
            seg_analysis['avg_coll_rate'] = seg_df['coll_rate'].mean()
            seg_analysis['coll_rate_trend'] = seg_df.groupby('Date')['coll_rate'].mean().diff().mean()

        # Interest revenue rate
        if 'InterestRevenue' in seg_df.columns and 'OpeningGBV' in seg_df.columns:
            seg_df['int_rate'] = seg_df['InterestRevenue'] / seg_df['OpeningGBV'].replace(0, np.nan)
            seg_analysis['avg_int_rate'] = seg_df['int_rate'].mean()

        # Writeoff rate
        if 'WO_DebtSold' in seg_df.columns and 'WO_Other' in seg_df.columns:
            seg_df['wo_rate'] = (seg_df['WO_DebtSold'] + seg_df['WO_Other']) / seg_df['OpeningGBV'].replace(0, np.nan)
            seg_analysis['avg_wo_rate'] = seg_df['wo_rate'].mean()

        # Coverage ratio
        if 'Total_Coverage_Ratio' in seg_df.columns:
            seg_analysis['avg_coverage'] = seg_df['Total_Coverage_Ratio'].mean()
            seg_analysis['coverage_trend'] = seg_df.groupby('Date')['Total_Coverage_Ratio'].mean().diff().mean()

        analysis[segment] = seg_analysis

    return analysis


def compare_to_budget_detailed(forecast_df, budget_comparison_file=None):
    """
    Compare forecast to budget and generate detailed variance analysis.

    If budget_comparison_file is provided, use it. Otherwise, generate from aggregation.
    """
    if budget_comparison_file and os.path.exists(budget_comparison_file):
        return pd.read_csv(budget_comparison_file)

    # Aggregate forecast by segment and month
    forecast_df['Date'] = pd.to_datetime(forecast_df['Date'])

    # Group by segment and date
    agg_metrics = {
        'ClosingGBV': 'sum',
        'Coll_Principal': 'sum',
        'Coll_Interest': 'sum',
        'InterestRevenue': 'sum',
        'WO_DebtSold': 'sum',
        'WO_Other': 'sum',
        'Net_Impairment': 'sum',
        'Gross_Impairment': 'sum',
        'ClosingNBV': 'sum'
    }

    # Only aggregate columns that exist
    existing_metrics = {k: v for k, v in agg_metrics.items() if k in forecast_df.columns}

    summary = forecast_df.groupby(['Segment', 'Date']).agg(existing_metrics).reset_index()

    # Calculate derived metrics
    if 'Coll_Principal' in summary.columns and 'Coll_Interest' in summary.columns:
        summary['Collections'] = summary['Coll_Principal'].abs() + summary['Coll_Interest'].abs()

    if 'InterestRevenue' in summary.columns:
        summary['Revenue'] = summary['InterestRevenue']

    return summary


def generate_iteration_report(iteration_num, methodology_file, forecast_summary,
                             variance_analysis, adjustments_made, rationale):
    """Generate a markdown report for this iteration."""
    report = f"""# Calibration Iteration {iteration_num}

## Date: {datetime.now().strftime('%Y-%m-%d %H:%M')}

## Methodology Used
File: `{methodology_file}`

## Adjustments Made in This Iteration
{adjustments_made}

## Rationale for Adjustments
{rationale}

## Variance Analysis Summary

### By Segment (Forecast Period Totals)
"""

    if variance_analysis is not None:
        # Summarize variances
        if isinstance(variance_analysis, pd.DataFrame):
            report += variance_analysis.to_markdown(index=False)

    report += f"""

## Key Observations

{generate_observations(variance_analysis)}

## Next Steps
Based on this analysis, the following adjustments should be considered for the next iteration:

{generate_next_steps(variance_analysis)}

---
"""
    return report


def generate_observations(variance_df):
    """Generate observations based on variance analysis."""
    if variance_df is None:
        return "- Unable to generate observations without variance data"

    observations = []

    # This will be populated based on actual variance data
    observations.append("- See variance analysis above for detailed metrics")

    return "\n".join(observations)


def generate_next_steps(variance_df):
    """Generate recommended next steps based on variance analysis."""
    if variance_df is None:
        return "- Run forecast and analyze results"

    return "- Review variance analysis and determine if further adjustments needed"


def save_iteration(iteration_num, methodology_df, methodology_file,
                  forecast_df, variance_df, report):
    """Save all iteration artifacts."""
    iter_dir = os.path.join(OUTPUT_DIR, f'iteration_{iteration_num}')
    os.makedirs(iter_dir, exist_ok=True)

    # Save methodology
    methodology_df.to_csv(os.path.join(iter_dir, 'Rate_Methodology.csv'), index=False)

    # Save forecast summary if available
    if forecast_df is not None:
        forecast_df.to_csv(os.path.join(iter_dir, 'Forecast_Summary.csv'), index=False)

    # Save variance analysis if available
    if variance_df is not None:
        variance_df.to_csv(os.path.join(iter_dir, 'Variance_Analysis.csv'), index=False)

    # Save report
    with open(os.path.join(iter_dir, 'ITERATION_REPORT.md'), 'w') as f:
        f.write(report)

    print(f"Iteration {iteration_num} saved to {iter_dir}")
    return iter_dir


if __name__ == '__main__':
    print("Iterative Calibration Analysis")
    print("=" * 50)

    # Analyze historical rates to understand the data
    print("\nAnalyzing historical rate patterns...")
    hist_analysis = analyze_historical_rates(FACT_RAW)

    print("\nHistorical Rate Analysis by Segment:")
    for segment, metrics in hist_analysis.items():
        print(f"\n{segment}:")
        for metric, value in metrics.items():
            if value is not None and not np.isnan(value):
                print(f"  {metric}: {value:.4f}")
