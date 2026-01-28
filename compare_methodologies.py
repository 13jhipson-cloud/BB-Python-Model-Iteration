#!/usr/bin/env python3
"""
Compare forecasts from different methodologies side by side.
"""

import pandas as pd
import numpy as np

def main():
    print("="*80)
    print("METHODOLOGY COMPARISON: ORIGINAL vs CURRENT")
    print("="*80)

    # Load both combined views
    print("\nLoading data...")
    current = pd.read_excel('output/Forecast_Transparency_Report.xlsx', sheet_name='6_Combined_View')
    original = pd.read_excel('output_original/Forecast_Transparency_Report.xlsx', sheet_name='6_Combined_View')

    # Focus on forecasts only
    current_fc = current[current['DataType'] == 'Forecast'].copy()
    original_fc = original[original['DataType'] == 'Forecast'].copy()

    # Key metrics to compare
    metrics = ['Total_Coverage_Ratio', 'Coll_Principal_Rate', 'InterestRevenue_Rate']

    # Get unique segment x cohort combinations
    combos = current_fc.groupby(['Segment', 'Cohort']).size().reset_index()[['Segment', 'Cohort']]

    print(f"\nFound {len(combos)} Segment × Cohort combinations")

    print("\n" + "="*80)
    print("TOTAL COVERAGE RATIO COMPARISON")
    print("="*80)

    for _, combo in combos.iterrows():
        segment = combo['Segment']
        cohort = combo['Cohort']

        curr = current_fc[(current_fc['Segment'] == segment) & (current_fc['Cohort'] == cohort)].sort_values('MOB')
        orig = original_fc[(original_fc['Segment'] == segment) & (original_fc['Cohort'] == cohort)].sort_values('MOB')

        print(f"\n{segment} | {cohort}")
        print("-"*60)
        print(f"{'MOB':>6} | {'Original':>12} | {'Current':>12} | {'Diff':>8} | {'Δ%':>8}")
        print("-"*60)

        for mob in curr['MOB'].unique():
            curr_val = curr[curr['MOB'] == mob]['Total_Coverage_Ratio'].values[0]
            orig_val = orig[orig['MOB'] == mob]['Total_Coverage_Ratio'].values[0] if mob in orig['MOB'].values else np.nan

            if not np.isnan(orig_val) and orig_val != 0:
                diff = curr_val - orig_val
                diff_pct = (curr_val - orig_val) / orig_val * 100
                print(f"{mob:>6} | {orig_val:>12.4f} | {curr_val:>12.4f} | {diff:>+8.4f} | {diff_pct:>+7.1f}%")
            else:
                print(f"{mob:>6} | {'N/A':>12} | {curr_val:>12.4f} | {'N/A':>8} | {'N/A':>8}")

    # Summary statistics
    print("\n" + "="*80)
    print("SUMMARY STATISTICS")
    print("="*80)

    # Merge on common keys
    merged = current_fc.merge(
        original_fc[['Segment', 'Cohort', 'MOB', 'Total_Coverage_Ratio']],
        on=['Segment', 'Cohort', 'MOB'],
        how='inner',
        suffixes=('_current', '_original')
    )

    merged['CR_Diff'] = merged['Total_Coverage_Ratio_current'] - merged['Total_Coverage_Ratio_original']
    merged['CR_Diff_Pct'] = merged['CR_Diff'] / merged['Total_Coverage_Ratio_original'] * 100

    print(f"\nTotal Coverage Ratio Differences (Current - Original):")
    print(f"  Mean Difference: {merged['CR_Diff'].mean():+.4f}")
    print(f"  Mean % Difference: {merged['CR_Diff_Pct'].mean():+.1f}%")
    print(f"  Min Difference: {merged['CR_Diff'].min():+.4f} ({merged['CR_Diff_Pct'].min():+.1f}%)")
    print(f"  Max Difference: {merged['CR_Diff'].max():+.4f} ({merged['CR_Diff_Pct'].max():+.1f}%)")

    # First forecast month only
    first_month = merged.groupby(['Segment', 'Cohort'])['MOB'].min().reset_index()
    first_month = first_month.merge(merged, on=['Segment', 'Cohort', 'MOB'])

    print(f"\nFirst Forecast Month Only:")
    print(f"  Mean Difference: {first_month['CR_Diff'].mean():+.4f}")
    print(f"  Mean % Difference: {first_month['CR_Diff_Pct'].mean():+.1f}%")

    # Last forecast month only
    last_month = merged.groupby(['Segment', 'Cohort'])['MOB'].max().reset_index()
    last_month = last_month.merge(merged, on=['Segment', 'Cohort', 'MOB'])

    print(f"\nLast Forecast Month Only:")
    print(f"  Mean Difference: {last_month['CR_Diff'].mean():+.4f}")
    print(f"  Mean % Difference: {last_month['CR_Diff_Pct'].mean():+.1f}%")

    print("\n" + "="*80)
    print("CONCLUSION")
    print("="*80)
    print("""
The current methodology uses ScaledCohortAvg with 1.2x multiplier for Total_Coverage_Ratio.
This causes COMPOUNDING because:
  1. Month 1: CR = historical_avg × 1.2
  2. Month 2: CR = rolling_avg(historical + month1) × 1.2
  3. Month 3: CR = rolling_avg(historical + month1 + month2) × 1.2

The 1.2x factor gets applied to already-inflated values, causing CR to grow exponentially.

The ORIGINAL methodology used:
  - CohortAvg (no multiplier) for mature cohorts (MOB 40+)
  - CohortTrend for mid-age cohorts (MOB 20-39)
  - DonorCohort for young cohorts (MOB 0-19)

This produced more sensible, stable forecasts that matched historical patterns.
""")

if __name__ == '__main__':
    main()
