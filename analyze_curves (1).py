#!/usr/bin/env python3
"""
Comprehensive Curve Analysis Script
Analyzes all forecast curves for each cohort x segment and identifies issues.
"""

import pandas as pd
import numpy as np
from openpyxl import load_workbook

def load_combined_view(xlsx_path='output/Forecast_Transparency_Report.xlsx'):
    """Load the combined view sheet from transparency report."""
    df = pd.read_excel(xlsx_path, sheet_name='6_Combined_View')
    return df

def analyze_metric_curve(df, segment, cohort, metric):
    """Analyze a single metric curve for a cohort x segment."""
    subset = df[(df['Segment'] == segment) & (df['Cohort'] == cohort)].copy()
    subset = subset.sort_values('MOB')

    actuals = subset[subset['DataType'] == 'Actual']
    forecasts = subset[subset['DataType'] == 'Forecast']

    if len(actuals) == 0 or len(forecasts) == 0:
        return None

    if metric not in actuals.columns:
        return None

    # Get last actual value
    last_actual_mob = actuals['MOB'].max()
    last_actual = actuals[actuals['MOB'] == last_actual_mob][metric].values[0]

    # Get first forecast value
    first_forecast_mob = forecasts['MOB'].min()
    first_forecast = forecasts[forecasts['MOB'] == first_forecast_mob][metric].values[0]

    # Calculate jump at transition
    if last_actual != 0:
        jump_pct = (first_forecast - last_actual) / abs(last_actual) * 100
    else:
        jump_pct = np.inf if first_forecast != 0 else 0

    # Calculate actuals stats
    actuals_mean = actuals[metric].mean()
    actuals_std = actuals[metric].std()
    actuals_min = actuals[metric].min()
    actuals_max = actuals[metric].max()

    # Calculate forecast stats
    forecast_mean = forecasts[metric].mean()
    forecast_std = forecasts[metric].std()
    forecast_min = forecasts[metric].min()
    forecast_max = forecasts[metric].max()

    # Calculate trend in actuals (simple linear regression)
    if len(actuals) >= 3:
        x = actuals['MOB'].values
        y = actuals[metric].values
        if not np.isnan(y).any() and np.std(x) > 0:
            slope = np.polyfit(x, y, 1)[0]
        else:
            slope = 0
    else:
        slope = 0

    # Flag issues
    issues = []

    # Issue 1: Large jump at transition (>20%)
    if abs(jump_pct) > 20 and last_actual != 0:
        issues.append(f"Large jump at transition: {jump_pct:+.1f}%")

    # Issue 2: Forecast mean very different from actuals mean (>30%)
    if actuals_mean != 0:
        mean_diff_pct = (forecast_mean - actuals_mean) / abs(actuals_mean) * 100
        if abs(mean_diff_pct) > 30:
            issues.append(f"Forecast mean differs from actuals: {mean_diff_pct:+.1f}%")

    # Issue 3: Forecast values outside historical range
    if forecast_max > actuals_max * 1.5 and actuals_max != 0:
        issues.append(f"Forecast max ({forecast_max:.4f}) > 1.5x actuals max ({actuals_max:.4f})")
    if forecast_min < actuals_min * 0.5 and actuals_min > 0:
        issues.append(f"Forecast min ({forecast_min:.4f}) < 0.5x actuals min ({actuals_min:.4f})")

    # Issue 4: Flat forecast when actuals show trend
    if abs(slope) > 0.0005 and forecast_std < 0.0001:
        issues.append(f"Flat forecast but actuals have trend (slope={slope:.6f})")

    # Issue 5: All zeros in forecast but not in actuals
    if forecasts[metric].sum() == 0 and actuals[metric].sum() != 0:
        issues.append("Forecast is all zeros but actuals are not")

    # Issue 6: All zeros in actuals but not in forecast
    if actuals[metric].sum() == 0 and forecasts[metric].sum() != 0:
        issues.append("Actuals are all zeros but forecast is not")

    return {
        'Segment': segment,
        'Cohort': cohort,
        'Metric': metric,
        'Last_Actual_MOB': last_actual_mob,
        'Last_Actual_Value': last_actual,
        'First_Forecast_MOB': first_forecast_mob,
        'First_Forecast_Value': first_forecast,
        'Jump_Pct': jump_pct,
        'Actuals_Mean': actuals_mean,
        'Actuals_Std': actuals_std,
        'Actuals_Min': actuals_min,
        'Actuals_Max': actuals_max,
        'Actuals_Trend_Slope': slope,
        'Forecast_Mean': forecast_mean,
        'Forecast_Std': forecast_std,
        'Forecast_Min': forecast_min,
        'Forecast_Max': forecast_max,
        'Issues': '; '.join(issues) if issues else 'OK',
        'Has_Issues': len(issues) > 0
    }


def main():
    print("="*80)
    print("COMPREHENSIVE FORECAST CURVE ANALYSIS")
    print("="*80)

    # Load data
    print("\nLoading data...")
    df = load_combined_view()

    # Get all unique segments and cohorts
    segments = df['Segment'].unique()
    cohorts = df['Cohort'].unique()

    print(f"Found {len(segments)} segments: {list(segments)}")
    print(f"Found {len(cohorts)} cohorts: {list(cohorts)}")

    # Metrics to analyze (rates)
    rate_metrics = [
        'Coll_Principal_Rate',
        'Coll_Interest_Rate',
        'InterestRevenue_Rate',
        'WO_DebtSold_Rate',
        'WO_Other_Rate',
        'Total_Coverage_Ratio'
    ]

    # Analyze each cohort x segment x metric
    results = []
    for segment in segments:
        for cohort in cohorts:
            # Check if this combination exists
            subset = df[(df['Segment'] == segment) & (df['Cohort'] == cohort)]
            if len(subset) == 0:
                continue

            for metric in rate_metrics:
                result = analyze_metric_curve(df, segment, cohort, metric)
                if result:
                    results.append(result)

    results_df = pd.DataFrame(results)

    # Print summary
    print("\n" + "="*80)
    print("SUMMARY")
    print("="*80)

    total_curves = len(results_df)
    curves_with_issues = results_df['Has_Issues'].sum()
    print(f"\nTotal curves analyzed: {total_curves}")
    print(f"Curves with issues: {curves_with_issues} ({curves_with_issues/total_curves*100:.1f}%)")
    print(f"Curves OK: {total_curves - curves_with_issues} ({(total_curves-curves_with_issues)/total_curves*100:.1f}%)")

    # Group by metric
    print("\n" + "-"*80)
    print("ISSUES BY METRIC")
    print("-"*80)

    for metric in rate_metrics:
        metric_df = results_df[results_df['Metric'] == metric]
        issues_count = metric_df['Has_Issues'].sum()
        total = len(metric_df)
        print(f"\n{metric}:")
        print(f"  Total curves: {total}, Issues: {issues_count} ({issues_count/total*100:.1f}% problematic)")

        if issues_count > 0:
            issue_rows = metric_df[metric_df['Has_Issues']]
            for _, row in issue_rows.iterrows():
                print(f"    [{row['Segment']}|{row['Cohort']}] {row['Issues']}")

    # Detailed analysis by segment x cohort
    print("\n" + "="*80)
    print("DETAILED ANALYSIS BY SEGMENT x COHORT")
    print("="*80)

    for segment in sorted(segments):
        segment_df = results_df[results_df['Segment'] == segment]
        cohorts_in_segment = segment_df['Cohort'].unique()

        print(f"\n{'='*80}")
        print(f"SEGMENT: {segment}")
        print(f"{'='*80}")

        for cohort in sorted(cohorts_in_segment):
            cohort_df = segment_df[segment_df['Cohort'] == cohort]
            issues_count = cohort_df['Has_Issues'].sum()

            print(f"\n  COHORT: {cohort}")
            print(f"  {'-'*60}")

            for _, row in cohort_df.iterrows():
                metric = row['Metric']
                status = "ISSUE" if row['Has_Issues'] else "OK"

                # Format values for display
                last_val = row['Last_Actual_Value']
                first_val = row['First_Forecast_Value']
                jump = row['Jump_Pct']

                # Create detailed output
                print(f"    {metric}:")
                print(f"      Last Actual (MOB {row['Last_Actual_MOB']:.0f}): {last_val:.6f}")
                print(f"      First Forecast (MOB {row['First_Forecast_MOB']:.0f}): {first_val:.6f}")
                if np.isfinite(jump):
                    print(f"      Transition Jump: {jump:+.1f}%")
                print(f"      Actuals Range: [{row['Actuals_Min']:.6f}, {row['Actuals_Max']:.6f}], Mean: {row['Actuals_Mean']:.6f}")
                print(f"      Forecast Range: [{row['Forecast_Min']:.6f}, {row['Forecast_Max']:.6f}], Mean: {row['Forecast_Mean']:.6f}")
                print(f"      Actuals Trend (slope): {row['Actuals_Trend_Slope']:.8f}")
                print(f"      Status: [{status}] {row['Issues']}")
                print()

    # Save to CSV
    output_path = 'output/curve_analysis_results.csv'
    results_df.to_csv(output_path, index=False)
    print(f"\nResults saved to: {output_path}")

    return results_df


if __name__ == '__main__':
    main()
