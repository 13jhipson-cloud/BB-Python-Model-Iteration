#!/usr/bin/env python3
"""
Auto-Calibration System for Backbook Forecast
==============================================

This script automatically iterates the forecast to match budget targets by:
1. Running the forecast model
2. Comparing outputs to the Budget consol file
3. Analyzing variances to identify which rate parameters to adjust
4. Updating the Rate_Methodology file with calibrated values
5. Re-running until variances are within tolerance (0.1 = 1 decimal place)

Usage:
    python auto_calibrate.py [--max-iterations 10] [--tolerance 0.1]

The script modifies Rate_Methodology.csv with calibrated values.
"""

import pandas as pd
import numpy as np
import os
import sys
import argparse
import logging
from datetime import datetime
from typing import Dict, List, Tuple, Optional
import shutil

# Import from backbook_forecast
from backbook_forecast import (
    run_backbook_forecast, load_fact_raw, load_rate_methodology,
    Config
)

# =============================================================================
# CONFIGURATION
# =============================================================================

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

# Segment mapping: Model segment -> Budget segment
SEGMENT_MAPPING = {
    'NON PRIME': 'Non Prime',
    'NRP-S': 'Near Prime Small',
    'NRP-M': 'Near Prime Medium',
    'NRP-L': 'Near Prime Medium',  # Combined
    'PRIME': 'Prime',
}

# Reverse mapping
BUDGET_TO_MODEL_SEGMENT = {v: k for k, v in SEGMENT_MAPPING.items()}
# Handle the NRP-M/NRP-L combination
BUDGET_TO_MODEL_SEGMENT['Near Prime Medium'] = ['NRP-M', 'NRP-L']

# Budget file row indices for each metric (0-indexed)
BUDGET_ROWS = {
    'Collections': {'Non Prime': 11, 'Near Prime Small': 12, 'Near Prime Medium': 13, 'Prime': 14, 'Total': 15},
    'ClosingGBV': {'Non Prime': 22, 'Near Prime Small': 23, 'Near Prime Medium': 24, 'Prime': 25, 'Total': 26},
    'ClosingNBV': {'Non Prime': 42, 'Near Prime Small': 43, 'Near Prime Medium': 44, 'Prime': 45, 'Total': 46},
    'Revenue': {'Non Prime': 62, 'Near Prime Small': 63, 'Near Prime Medium': 64, 'Prime': 65, 'Total': 66},
    'GrossImpairment': {'Non Prime': 73, 'Near Prime Small': 74, 'Near Prime Medium': 75, 'Prime': 76, 'Total': 77},
    'DebtSaleGain': {'Non Prime': 105, 'Near Prime Small': 106, 'Near Prime Medium': 107, 'Prime': 108, 'Total': 109},
    'NetImpairment': {'Non Prime': 121, 'Near Prime Small': 122, 'Near Prime Medium': 123, 'Prime': 124, 'Total': 125},
}

# Mapping from variance metric to rate methodology metric to adjust
# This tells us which rates affect which outputs
METRIC_TO_RATE_MAPPING = {
    'Collections': ['Coll_Principal', 'Coll_Interest'],  # Collections variance -> adjust collection rates
    'ClosingGBV': ['Coll_Principal', 'Coll_Interest', 'WO_Other', 'WO_DebtSold'],  # GBV affected by collections and writeoffs
    'Revenue': ['InterestRevenue'],  # Revenue variance -> adjust interest revenue rate
    'GrossImpairment': ['Total_Coverage_Ratio', 'WO_Other'],  # Gross impairment driven by coverage and writeoffs
    'NetImpairment': ['Total_Coverage_Ratio', 'WO_Other', 'WO_DebtSold'],  # Net impairment = gross + debt sale
    'ClosingNBV': ['Total_Coverage_Ratio'],  # NBV = GBV - Provision, so coverage ratio is key
}

# =============================================================================
# BUDGET LOADING
# =============================================================================

def load_budget(budget_path: str) -> pd.DataFrame:
    """Load budget data from Excel and return in long format."""
    logger.info(f"Loading budget from: {budget_path}")

    xl = pd.ExcelFile(budget_path)
    df_raw = pd.read_excel(xl, sheet_name='P&L analysis - BB', header=None)

    # Extract dates from row 2, columns 3 onwards (skip the Excel date number)
    dates = []
    for val in df_raw.iloc[2, 3:].values:
        if isinstance(val, datetime):
            dates.append(pd.Timestamp(val))
        elif pd.notna(val):
            try:
                dates.append(pd.to_datetime(val))
            except:
                continue

    records = []
    for metric, segment_rows in BUDGET_ROWS.items():
        for segment, row_idx in segment_rows.items():
            values = df_raw.iloc[row_idx, 3:3+len(dates)].values
            for date, value in zip(dates, values):
                if pd.notna(value):
                    try:
                        records.append({
                            'Date': date,
                            'Segment': segment,
                            'Metric': metric,
                            'Budget_Value': float(value)
                        })
                    except (ValueError, TypeError):
                        continue

    budget_df = pd.DataFrame(records)
    logger.info(f"  Loaded {len(budget_df)} budget data points")
    return budget_df


# =============================================================================
# FORECAST RUNNING
# =============================================================================

def run_forecast_and_get_outputs(
    fact_raw_path: str,
    methodology_path: str,
    output_dir: str = 'calibration_output',
    forecast_months: int = 12
) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """
    Run the forecast model and return the outputs as DataFrames.

    Returns:
        tuple: (summary_df, impairment_df)
    """
    logger.info("Running forecast model...")

    # Run forecast
    result = run_backbook_forecast(
        fact_raw_path=fact_raw_path,
        methodology_path=methodology_path,
        debt_sale_path=None,
        output_dir=output_dir,
        max_months=forecast_months
    )

    # Load the outputs
    summary_path = os.path.join(output_dir, 'Forecast_Summary.xlsx')
    impairment_path = os.path.join(output_dir, 'Impairment_Analysis.xlsx')

    summary_df = pd.read_excel(summary_path, sheet_name='Summary')
    impairment_df = pd.read_excel(impairment_path, sheet_name='Impairment_Detail')

    summary_df['ForecastMonth'] = pd.to_datetime(summary_df['ForecastMonth'])
    impairment_df['ForecastMonth'] = pd.to_datetime(impairment_df['ForecastMonth'])

    logger.info(f"  Forecast completed: {len(summary_df)} rows in summary")
    return summary_df, impairment_df


def aggregate_forecast_for_comparison(
    summary_df: pd.DataFrame,
    impairment_df: pd.DataFrame
) -> pd.DataFrame:
    """
    Aggregate forecast outputs by budget segment and month.

    Returns DataFrame with: [Date, Segment, Metric, Forecast_Value]
    """
    # Map to budget segments
    summary_df['BudgetSegment'] = summary_df['Segment'].map(SEGMENT_MAPPING)
    impairment_df['BudgetSegment'] = impairment_df['Segment'].map(SEGMENT_MAPPING)

    # Aggregate summary
    agg_df = summary_df.groupby(['ForecastMonth', 'BudgetSegment']).agg({
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

    # Aggregate impairment
    imp_agg = impairment_df.groupby(['ForecastMonth', 'BudgetSegment']).agg({
        'Gross_Impairment_ExcludingDS': 'sum',
        'Debt_Sale_Impact': 'sum',
    }).reset_index()

    # Merge
    agg_df = agg_df.merge(
        imp_agg[['ForecastMonth', 'BudgetSegment', 'Gross_Impairment_ExcludingDS', 'Debt_Sale_Impact']],
        on=['ForecastMonth', 'BudgetSegment'],
        how='left'
    )

    # Calculate derived metrics (matching budget conventions)
    agg_df['Collections'] = -(agg_df['Coll_Principal'] + agg_df['Coll_Interest'])
    agg_df['Revenue'] = agg_df['InterestRevenue']
    agg_df['GrossImpairment'] = agg_df['Gross_Impairment_ExcludingDS']
    # With new sign convention, Debt_Sale_Impact already represents gain directly
    agg_df['DebtSaleGain'] = agg_df['Debt_Sale_Impact']
    agg_df['NetImpairment'] = agg_df['Net_Impairment']

    # Reshape to long format
    records = []
    metrics = ['Collections', 'ClosingGBV', 'ClosingNBV', 'Revenue',
               'GrossImpairment', 'DebtSaleGain', 'NetImpairment']

    for _, row in agg_df.iterrows():
        for metric in metrics:
            records.append({
                'Date': row['ForecastMonth'],
                'Segment': row['BudgetSegment'],
                'Metric': metric,
                'Forecast_Value': row[metric]
            })

    # Add totals
    total_df = agg_df.groupby('ForecastMonth')[metrics].sum().reset_index()
    for _, row in total_df.iterrows():
        for metric in metrics:
            records.append({
                'Date': row['ForecastMonth'],
                'Segment': 'Total',
                'Metric': metric,
                'Forecast_Value': row[metric]
            })

    return pd.DataFrame(records)


# =============================================================================
# VARIANCE ANALYSIS
# =============================================================================

def calculate_variances(
    budget_df: pd.DataFrame,
    forecast_df: pd.DataFrame
) -> pd.DataFrame:
    """
    Calculate variances between budget and forecast.

    Returns DataFrame with variance analysis.
    """
    # Find common dates
    budget_dates = set(budget_df['Date'].unique())
    forecast_dates = set(forecast_df['Date'].unique())
    common_dates = budget_dates & forecast_dates

    if not common_dates:
        logger.warning("No common dates between budget and forecast!")
        return pd.DataFrame()

    logger.info(f"  Comparing {len(common_dates)} months")

    budget_filtered = budget_df[budget_df['Date'].isin(common_dates)]
    forecast_filtered = forecast_df[forecast_df['Date'].isin(common_dates)]

    # Merge
    merged = pd.merge(
        budget_filtered,
        forecast_filtered,
        on=['Date', 'Segment', 'Metric'],
        how='outer'
    )

    # Calculate variances
    merged['Variance'] = merged['Forecast_Value'] - merged['Budget_Value']
    merged['Variance_Pct'] = np.where(
        merged['Budget_Value'] != 0,
        (merged['Forecast_Value'] / merged['Budget_Value'] - 1) * 100,
        np.nan
    )
    merged['Abs_Variance'] = np.abs(merged['Variance'])

    return merged


def identify_largest_variances(variance_df: pd.DataFrame, top_n: int = 5) -> pd.DataFrame:
    """
    Identify the largest variances that need adjustment.

    Returns a DataFrame summarizing the top variances by segment and metric.
    """
    # Group by segment and metric to get total variance
    summary = variance_df.groupby(['Segment', 'Metric']).agg({
        'Budget_Value': 'sum',
        'Forecast_Value': 'sum',
        'Variance': 'sum',
        'Abs_Variance': 'sum'
    }).reset_index()

    summary['Variance_Pct'] = np.where(
        summary['Budget_Value'] != 0,
        (summary['Forecast_Value'] / summary['Budget_Value'] - 1) * 100,
        np.nan
    )

    # Sort by absolute variance percentage
    summary['Abs_Variance_Pct'] = np.abs(summary['Variance_Pct'])
    summary = summary.sort_values('Abs_Variance_Pct', ascending=False)

    # Filter out totals for adjustment purposes (adjust segments, not totals)
    segment_summary = summary[summary['Segment'] != 'Total'].head(top_n)

    return segment_summary


# =============================================================================
# RATE METHODOLOGY ADJUSTMENT
# =============================================================================

def load_methodology(path: str) -> pd.DataFrame:
    """Load rate methodology CSV."""
    return pd.read_csv(path)


def save_methodology(df: pd.DataFrame, path: str, backup: bool = True):
    """Save rate methodology CSV, optionally creating a backup."""
    if backup:
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        backup_path = path.replace('.csv', f'_backup_{timestamp}.csv')
        shutil.copy(path, backup_path)
        logger.info(f"  Backup saved to: {backup_path}")

    df.to_csv(path, index=False)
    logger.info(f"  Methodology saved to: {path}")


def calculate_adjustment_factor(
    budget_value: float,
    forecast_value: float,
    damping: float = 0.5
) -> float:
    """
    Calculate the adjustment factor needed to move forecast toward budget.

    Args:
        budget_value: Target budget value
        forecast_value: Current forecast value
        damping: Damping factor (0-1) to prevent overshooting. Default 0.5 = 50% correction per iteration.

    Returns:
        Multiplier to apply to the rate (e.g., 1.1 means increase by 10%)
    """
    if forecast_value == 0 or budget_value == 0:
        return 1.0

    # Raw adjustment ratio
    raw_ratio = budget_value / forecast_value

    # Apply damping to avoid overshooting
    # If raw_ratio is 1.2 (need 20% increase) and damping is 0.5:
    # adjusted_ratio = 1 + (1.2 - 1) * 0.5 = 1.1 (only 10% increase)
    adjusted_ratio = 1 + (raw_ratio - 1) * damping

    # Clamp to reasonable bounds
    return max(0.5, min(2.0, adjusted_ratio))


def adjust_methodology_for_variance(
    methodology_df: pd.DataFrame,
    variance_summary: pd.DataFrame,
    damping: float = 0.5
) -> pd.DataFrame:
    """
    Adjust rate methodology based on variance analysis.

    For each large variance:
    1. Identify which rate metric(s) affect it
    2. Identify which methodology rules apply to that segment/metric
    3. Adjust the Param1 value (scaling factor) or add a multiplier

    Returns: Modified methodology DataFrame
    """
    logger.info("Adjusting methodology based on variances...")

    meth = methodology_df.copy()

    for _, var_row in variance_summary.iterrows():
        budget_segment = var_row['Segment']
        output_metric = var_row['Metric']
        variance_pct = var_row['Variance_Pct']
        budget_val = var_row['Budget_Value']
        forecast_val = var_row['Forecast_Value']

        # Skip if variance is small (within 1%)
        if abs(variance_pct) < 1:
            continue

        logger.info(f"  {budget_segment} {output_metric}: {variance_pct:+.1f}% variance")

        # Get the model segments for this budget segment
        model_segments = BUDGET_TO_MODEL_SEGMENT.get(budget_segment, [])
        if isinstance(model_segments, str):
            model_segments = [model_segments]

        # Get which rate metrics affect this output metric
        rate_metrics = METRIC_TO_RATE_MAPPING.get(output_metric, [])

        for rate_metric in rate_metrics:
            # Calculate adjustment
            adj_factor = calculate_adjustment_factor(budget_val, forecast_val, damping)

            # If forecast > budget, we need to reduce rates (adj_factor < 1)
            # If forecast < budget, we need to increase rates (adj_factor > 1)

            # Handle special logic per metric
            if output_metric == 'Collections' and rate_metric in ['Coll_Principal', 'Coll_Interest']:
                # Collections: if forecast collections < budget, increase collection rates
                # Note: Collections in model are negative but we compare absolute values
                pass

            elif output_metric in ['GrossImpairment', 'NetImpairment'] and rate_metric == 'Total_Coverage_Ratio':
                # Impairment: if forecast > budget (too much impairment), reduce coverage ratio
                # (lower coverage = lower provision movement = lower impairment charge)
                pass

            # Find matching rules to adjust
            for model_seg in model_segments:
                mask = (
                    ((meth['Segment'] == model_seg) | (meth['Segment'] == 'ALL')) &
                    (meth['Metric'] == rate_metric)
                )

                matching_rules = meth[mask]

                if len(matching_rules) > 0:
                    # For Manual approach, adjust Param1 directly
                    # For CohortAvg/CohortTrend, we can add a scaling multiplier via Param2

                    for idx in matching_rules.index:
                        approach = meth.loc[idx, 'Approach']
                        current_param1 = meth.loc[idx, 'Param1']

                        if approach == 'Manual' and pd.notna(current_param1):
                            # Adjust Manual value directly
                            try:
                                new_val = float(current_param1) * adj_factor
                                meth.loc[idx, 'Param1'] = round(new_val, 6)
                                logger.info(f"    Adjusted {model_seg}/{rate_metric} Manual: {current_param1} -> {new_val:.6f}")
                            except (ValueError, TypeError):
                                pass

                        elif approach in ['CohortAvg', 'CohortTrend']:
                            # For averaging approaches, set Param2 as a multiplier
                            # This requires code change to support Param2 as scaling factor
                            # For now, log the suggested adjustment
                            logger.info(f"    Suggest scaling {model_seg}/{rate_metric} {approach} by {adj_factor:.3f}")

    return meth


def add_scaling_overrides(
    methodology_df: pd.DataFrame,
    variance_summary: pd.DataFrame,
    damping: float = 0.5
) -> pd.DataFrame:
    """
    Add Manual override rules for specific segments that need adjustment.

    This creates new rules with higher specificity that will override the
    general CohortAvg rules.
    """
    logger.info("Adding scaling override rules...")

    meth = methodology_df.copy()
    new_rules = []

    for _, var_row in variance_summary.iterrows():
        budget_segment = var_row['Segment']
        output_metric = var_row['Metric']
        variance_pct = var_row['Variance_Pct']
        budget_val = var_row['Budget_Value']
        forecast_val = var_row['Forecast_Value']

        # Skip small variances
        if abs(variance_pct) < 2:  # 2% threshold for adding overrides
            continue

        # Get model segments
        model_segments = BUDGET_TO_MODEL_SEGMENT.get(budget_segment, [])
        if isinstance(model_segments, str):
            model_segments = [model_segments]

        # Get rate metrics
        rate_metrics = METRIC_TO_RATE_MAPPING.get(output_metric, [])

        adj_factor = calculate_adjustment_factor(budget_val, forecast_val, damping)

        # Only add override for the primary driving metric
        primary_metric = rate_metrics[0] if rate_metrics else None

        if primary_metric and primary_metric in ['Coll_Principal', 'Coll_Interest', 'Total_Coverage_Ratio']:
            for model_seg in model_segments:
                # Check if we already have a ScaledCohortAvg rule
                existing = meth[
                    (meth['Segment'] == model_seg) &
                    (meth['Metric'] == primary_metric) &
                    (meth['Approach'] == 'ScaledCohortAvg')
                ]

                if len(existing) > 0:
                    # Update existing rule
                    for idx in existing.index:
                        current_scale = float(meth.loc[idx, 'Param2'] or 1.0)
                        new_scale = current_scale * adj_factor
                        meth.loc[idx, 'Param2'] = round(new_scale, 4)
                        logger.info(f"  Updated {model_seg}/{primary_metric} scale: {current_scale:.4f} -> {new_scale:.4f}")
                else:
                    # Add new ScaledCohortAvg rule
                    new_rule = {
                        'Segment': model_seg,
                        'Cohort': 'ALL',
                        'Metric': primary_metric,
                        'MOB_Start': 0,
                        'MOB_End': 999,
                        'Approach': 'ScaledCohortAvg',
                        'Param1': 6,  # Rolling periods
                        'Param2': round(adj_factor, 4),  # Scale factor
                        'Explanation': f'Auto-calibrated scale factor for {budget_segment} {output_metric}'
                    }
                    new_rules.append(new_rule)
                    logger.info(f"  Added ScaledCohortAvg rule: {model_seg}/{primary_metric} scale={adj_factor:.4f}")

    if new_rules:
        new_df = pd.DataFrame(new_rules)
        meth = pd.concat([meth, new_df], ignore_index=True)

    return meth


# =============================================================================
# CONVERGENCE CHECK
# =============================================================================

def check_convergence(variance_df: pd.DataFrame, tolerance: float = 0.1) -> Tuple[bool, float]:
    """
    Check if all variances are within tolerance.

    Args:
        variance_df: Variance analysis DataFrame
        tolerance: Acceptable variance percentage (0.1 = ±10%)

    Returns:
        tuple: (converged: bool, max_variance_pct: float)
    """
    # Calculate summary variance by segment (excluding Total)
    segment_variance = variance_df[variance_df['Segment'] != 'Total'].groupby(['Segment', 'Metric']).agg({
        'Budget_Value': 'sum',
        'Forecast_Value': 'sum'
    }).reset_index()

    segment_variance['Variance_Pct'] = np.where(
        segment_variance['Budget_Value'] != 0,
        (segment_variance['Forecast_Value'] / segment_variance['Budget_Value'] - 1) * 100,
        0
    )

    max_variance = segment_variance['Variance_Pct'].abs().max()
    converged = max_variance <= tolerance * 100  # tolerance is in decimal, variance is in percent

    return converged, max_variance


# =============================================================================
# MAIN CALIBRATION LOOP
# =============================================================================

def run_calibration(
    fact_raw_path: str = 'Fact_Raw.xlsx',
    methodology_path: str = 'sample_data/Rate_Methodology_v6_Simplified.csv',
    budget_path: str = 'Budget consol file.xlsx',
    output_dir: str = 'calibration_output',
    max_iterations: int = 10,
    tolerance: float = 0.1,
    damping: float = 0.5,
    forecast_months: int = 12
):
    """
    Main calibration loop.

    Args:
        fact_raw_path: Path to Fact_Raw data
        methodology_path: Path to Rate_Methodology CSV (will be modified)
        budget_path: Path to Budget consol file
        output_dir: Directory for forecast outputs
        max_iterations: Maximum calibration iterations
        tolerance: Convergence tolerance (0.1 = within 10%)
        damping: Damping factor for adjustments (0-1)
        forecast_months: Number of months to forecast
    """
    print("=" * 80)
    print("AUTO-CALIBRATION SYSTEM")
    print("=" * 80)
    print(f"Fact Raw: {fact_raw_path}")
    print(f"Methodology: {methodology_path}")
    print(f"Budget: {budget_path}")
    print(f"Max iterations: {max_iterations}")
    print(f"Tolerance: {tolerance * 100:.1f}%")
    print("=" * 80)

    # Load budget once
    budget_df = load_budget(budget_path)

    # Create output directory
    os.makedirs(output_dir, exist_ok=True)

    # Iteration history
    history = []

    for iteration in range(1, max_iterations + 1):
        print(f"\n{'='*40}")
        print(f"ITERATION {iteration}")
        print(f"{'='*40}")

        # Step 1: Run forecast
        try:
            summary_df, impairment_df = run_forecast_and_get_outputs(
                fact_raw_path, methodology_path, output_dir, forecast_months
            )
        except Exception as e:
            logger.error(f"Forecast failed: {e}")
            break

        # Step 2: Aggregate for comparison
        forecast_df = aggregate_forecast_for_comparison(summary_df, impairment_df)

        # Step 3: Calculate variances
        variance_df = calculate_variances(budget_df, forecast_df)

        if len(variance_df) == 0:
            logger.error("No variance data - check date alignment")
            break

        # Step 4: Check convergence
        converged, max_variance = check_convergence(variance_df, tolerance)

        history.append({
            'iteration': iteration,
            'max_variance_pct': max_variance,
            'converged': converged
        })

        print(f"\nMax variance: {max_variance:.2f}%")

        if converged:
            print(f"\n{'='*40}")
            print(f"CONVERGED after {iteration} iterations!")
            print(f"All variances within {tolerance*100:.1f}%")
            print(f"{'='*40}")
            break

        # Step 5: Identify largest variances
        top_variances = identify_largest_variances(variance_df, top_n=5)

        print("\nTop variances to address:")
        for _, row in top_variances.iterrows():
            print(f"  {row['Segment']:20} {row['Metric']:15} {row['Variance_Pct']:+7.2f}%")

        # Step 6: Adjust methodology
        methodology_df = load_methodology(methodology_path)

        # Apply adjustments
        adjusted_meth = adjust_methodology_for_variance(
            methodology_df, top_variances, damping
        )

        # Optionally add scaling overrides
        adjusted_meth = add_scaling_overrides(
            adjusted_meth, top_variances, damping
        )

        # Step 7: Save updated methodology
        save_methodology(adjusted_meth, methodology_path, backup=(iteration == 1))

        print(f"\nMethodology updated for iteration {iteration + 1}")

    else:
        print(f"\n{'='*40}")
        print(f"Did not converge within {max_iterations} iterations")
        print(f"Final max variance: {max_variance:.2f}%")
        print(f"{'='*40}")

    # Print iteration history
    print("\nIteration History:")
    print("-" * 40)
    for h in history:
        status = "✓ Converged" if h['converged'] else ""
        print(f"  Iteration {h['iteration']}: Max variance {h['max_variance_pct']:.2f}% {status}")

    # Save final variance report
    variance_df.to_csv(os.path.join(output_dir, 'final_variance_analysis.csv'), index=False)
    print(f"\nFinal variance analysis saved to: {output_dir}/final_variance_analysis.csv")

    return history


# =============================================================================
# CLI
# =============================================================================

def main():
    parser = argparse.ArgumentParser(
        description='Auto-calibrate forecast model to match budget targets'
    )
    parser.add_argument('--fact-raw', default='Fact_Raw.xlsx',
                        help='Path to Fact_Raw data file')
    parser.add_argument('--methodology', '-m', default='sample_data/Rate_Methodology_v6_Simplified.csv',
                        help='Path to Rate_Methodology CSV')
    parser.add_argument('--budget', '-b', default='Budget consol file.xlsx',
                        help='Path to Budget consol file')
    parser.add_argument('--output', '-o', default='calibration_output',
                        help='Output directory')
    parser.add_argument('--max-iterations', type=int, default=10,
                        help='Maximum calibration iterations')
    parser.add_argument('--tolerance', type=float, default=0.1,
                        help='Convergence tolerance (0.1 = 10%%)')
    parser.add_argument('--damping', type=float, default=0.5,
                        help='Adjustment damping factor (0-1)')
    parser.add_argument('--months', type=int, default=12,
                        help='Number of months to forecast')

    args = parser.parse_args()

    run_calibration(
        fact_raw_path=args.fact_raw,
        methodology_path=args.methodology,
        budget_path=args.budget,
        output_dir=args.output,
        max_iterations=args.max_iterations,
        tolerance=args.tolerance,
        damping=args.damping,
        forecast_months=args.months
    )


if __name__ == '__main__':
    main()
