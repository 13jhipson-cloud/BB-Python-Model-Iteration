#!/usr/bin/env python3
"""
Iterative Calibration Script
Adjusts Rate_Methodology to match budget targets by applying scaling factors.
"""

import pandas as pd
import numpy as np
import logging
from backbook_forecast import run_backbook_forecast
from auto_calibrate import load_budget, SEGMENT_MAPPING

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# Mapping from budget segments to model segments
BUDGET_TO_MODEL_SEGMENTS = {
    'Non Prime': ['NON PRIME'],
    'Near Prime Small': ['NRP-S'],
    'Near Prime Medium': ['NRP-M', 'NRP-L'],
    'Prime': ['PRIME']
}

def run_forecast_and_compare(methodology_path, budget_df, output_dir='calibration_iter'):
    """Run forecast and compare to budget, returning variance metrics."""

    # Run forecast
    result = run_backbook_forecast(
        fact_raw_path='Fact_Raw.xlsx',
        methodology_path=methodology_path,
        debt_sale_path=None,
        output_dir=output_dir,
        max_months=12
    )

    # Load results
    summary = pd.read_excel(f'{output_dir}/Forecast_Summary.xlsx', sheet_name='Summary')
    summary['ForecastMonth'] = pd.to_datetime(summary['ForecastMonth'])
    summary['BudgetSegment'] = summary['Segment'].map(SEGMENT_MAPPING)

    imp = pd.read_excel(f'{output_dir}/Impairment_Analysis.xlsx', sheet_name='Impairment_Detail')
    imp['ForecastMonth'] = pd.to_datetime(imp['ForecastMonth'])
    imp['BudgetSegment'] = imp['Segment'].map(SEGMENT_MAPPING)

    # Filter to budget months
    budget_months = budget_df[budget_df['Date'] >= '2025-11-01']['Date'].unique()
    summary_aligned = summary[summary['ForecastMonth'].isin(budget_months)]
    imp_aligned = imp[imp['ForecastMonth'].isin(budget_months)]

    # Calculate variances
    variances = {}

    for seg in ['Non Prime', 'Near Prime Small', 'Near Prime Medium', 'Prime']:
        seg_summary = summary_aligned[summary_aligned['BudgetSegment'] == seg]
        seg_imp = imp_aligned[imp_aligned['BudgetSegment'] == seg]

        # Model values
        model_coll = -(seg_summary['Coll_Principal'].sum() + seg_summary['Coll_Interest'].sum())
        model_rev = seg_summary['InterestRevenue'].sum()
        model_gbv = seg_summary.groupby('ForecastMonth')['ClosingGBV'].sum().iloc[-1] if len(seg_summary) > 0 else 0
        model_imp = seg_imp['Net_Impairment'].sum() if len(seg_imp) > 0 else 0

        # Budget values
        budget_coll = budget_df[(budget_df['Metric'] == 'Collections') &
                                (budget_df['Segment'] == seg) &
                                (budget_df['Date'].isin(budget_months))]['Budget_Value'].sum()
        budget_rev = budget_df[(budget_df['Metric'] == 'Revenue') &
                               (budget_df['Segment'] == seg) &
                               (budget_df['Date'].isin(budget_months))]['Budget_Value'].sum()
        budget_gbv = budget_df[(budget_df['Metric'] == 'ClosingGBV') &
                               (budget_df['Segment'] == seg) &
                               (budget_df['Date'] == '2026-09-30')]['Budget_Value'].sum()
        budget_imp = budget_df[(budget_df['Metric'] == 'NetImpairment') &
                               (budget_df['Segment'] == seg) &
                               (budget_df['Date'].isin(budget_months))]['Budget_Value'].sum()

        variances[seg] = {
            'Collections': {'model': model_coll, 'budget': budget_coll,
                           'var_pct': (model_coll/budget_coll - 1)*100 if budget_coll else 0},
            'Revenue': {'model': model_rev, 'budget': budget_rev,
                       'var_pct': (model_rev/budget_rev - 1)*100 if budget_rev else 0},
            'ClosingGBV': {'model': model_gbv, 'budget': budget_gbv,
                          'var_pct': (model_gbv/budget_gbv - 1)*100 if budget_gbv else 0},
            'NetImpairment': {'model': model_imp, 'budget': budget_imp,
                             'var_pct': (model_imp/budget_imp - 1)*100 if budget_imp else 0}
        }

    return variances, summary, imp


def calculate_scaling_factors(variances, damping=0.5):
    """Calculate scaling factors to apply to methodology, with damping."""

    factors = {}

    for seg, metrics in variances.items():
        factors[seg] = {}

        for metric, vals in metrics.items():
            if vals['model'] != 0 and vals['budget'] != 0:
                raw_scale = vals['budget'] / vals['model']
                # Apply damping to avoid overshooting
                damped_scale = 1 + (raw_scale - 1) * damping
                factors[seg][metric] = damped_scale
            else:
                factors[seg][metric] = 1.0

    return factors


def apply_scaling_to_methodology(methodology_path, factors, output_path):
    """Apply scaling factors to methodology file."""

    meth = pd.read_csv(methodology_path)

    # Add a ScaleFactor column if not exists
    if 'ScaleFactor' not in meth.columns:
        meth['ScaleFactor'] = 1.0

    # Map budget segments to model segments
    for budget_seg, model_segs in BUDGET_TO_MODEL_SEGMENTS.items():
        if budget_seg not in factors:
            continue

        seg_factors = factors[budget_seg]

        for model_seg in model_segs:
            # Apply collection scaling
            coll_scale = seg_factors.get('Collections', 1.0)
            mask = (meth['Segment'].isin([model_seg, 'ALL'])) & \
                   (meth['Metric'].isin(['Coll_Principal', 'Coll_Interest']))
            if model_seg != 'ALL':
                mask = (meth['Segment'] == model_seg) & \
                       (meth['Metric'].isin(['Coll_Principal', 'Coll_Interest']))
            meth.loc[mask, 'ScaleFactor'] = meth.loc[mask, 'ScaleFactor'] * coll_scale

            # Apply revenue scaling
            rev_scale = seg_factors.get('Revenue', 1.0)
            mask = (meth['Segment'] == model_seg) & (meth['Metric'] == 'InterestRevenue')
            meth.loc[mask, 'ScaleFactor'] = meth.loc[mask, 'ScaleFactor'] * rev_scale

            # Apply coverage ratio scaling (inverse of impairment scale for releases)
            imp_scale = seg_factors.get('NetImpairment', 1.0)
            # For impairment, we need to adjust coverage ratios
            # If model shows charge and budget shows release, we need lower coverage
            mask = (meth['Segment'] == model_seg) & (meth['Metric'] == 'Total_Coverage_Ratio')
            if imp_scale < 0:  # Sign flip needed
                # Need to reduce coverage ratios significantly
                meth.loc[mask, 'ScaleFactor'] = meth.loc[mask, 'ScaleFactor'] * 0.5
            else:
                meth.loc[mask, 'ScaleFactor'] = meth.loc[mask, 'ScaleFactor'] * min(imp_scale, 2.0)

    # For ALL segment rules, apply average scaling
    avg_coll_scale = np.mean([factors[s].get('Collections', 1.0) for s in factors])
    avg_rev_scale = np.mean([factors[s].get('Revenue', 1.0) for s in factors])

    mask = (meth['Segment'] == 'ALL') & (meth['Metric'].isin(['Coll_Principal', 'Coll_Interest']))
    meth.loc[mask, 'ScaleFactor'] = meth.loc[mask, 'ScaleFactor'] * avg_coll_scale

    mask = (meth['Segment'] == 'ALL') & (meth['Metric'] == 'InterestRevenue')
    meth.loc[mask, 'ScaleFactor'] = meth.loc[mask, 'ScaleFactor'] * avg_rev_scale

    meth.to_csv(output_path, index=False)
    return meth


def print_variance_summary(variances, iteration):
    """Print formatted variance summary."""

    print(f"\n{'='*100}")
    print(f"ITERATION {iteration} - Variance Summary")
    print(f"{'='*100}")

    metrics = ['Collections', 'Revenue', 'ClosingGBV', 'NetImpairment']

    for metric in metrics:
        print(f"\n{metric}:")
        print("-" * 80)
        total_model = 0
        total_budget = 0
        for seg in ['Non Prime', 'Near Prime Small', 'Near Prime Medium', 'Prime']:
            v = variances[seg][metric]
            print(f"  {seg:20} Model: {v['model']:>14,.0f}  Budget: {v['budget']:>14,.0f}  Var: {v['var_pct']:>+7.1f}%")
            total_model += v['model']
            total_budget += v['budget']
        total_var = (total_model/total_budget - 1)*100 if total_budget else 0
        print(f"  {'TOTAL':20} Model: {total_model:>14,.0f}  Budget: {total_budget:>14,.0f}  Var: {total_var:>+7.1f}%")


def run_iterative_calibration(max_iterations=10, tolerance=0.05, damping=0.3):
    """Main calibration loop."""

    # Load budget
    budget_df = load_budget('Budget consol file.xlsx')

    # Start with original methodology
    import shutil
    shutil.copy('pre_calibration_baseline/Rate_Methodology_ORIGINAL.csv',
                'sample_data/Rate_Methodology_CALIBRATING.csv')

    methodology_path = 'sample_data/Rate_Methodology_CALIBRATING.csv'

    best_variances = None
    best_iteration = 0
    best_max_var = float('inf')

    for iteration in range(1, max_iterations + 1):
        logger.info(f"Starting iteration {iteration}")

        # Run forecast and compare
        variances, summary, imp = run_forecast_and_compare(
            methodology_path, budget_df, f'calibration_iter_{iteration}'
        )

        # Print summary
        print_variance_summary(variances, iteration)

        # Calculate max variance (excluding impairment sign issues)
        max_var = 0
        for seg, metrics in variances.items():
            for metric, vals in metrics.items():
                if metric != 'NetImpairment':  # Exclude impairment for now
                    max_var = max(max_var, abs(vals['var_pct']))

        logger.info(f"Iteration {iteration}: Max variance = {max_var:.1f}%")

        # Track best result
        if max_var < best_max_var:
            best_max_var = max_var
            best_variances = variances
            best_iteration = iteration
            # Save best methodology
            shutil.copy(methodology_path, 'sample_data/Rate_Methodology_BEST.csv')

        # Check convergence
        if max_var < tolerance * 100:
            logger.info(f"Converged at iteration {iteration} with max variance {max_var:.1f}%")
            break

        # Calculate and apply scaling factors
        factors = calculate_scaling_factors(variances, damping=damping)

        logger.info(f"Applying scaling factors with damping={damping}")
        for seg, seg_factors in factors.items():
            logger.info(f"  {seg}: Coll={seg_factors.get('Collections', 1):.3f}, Rev={seg_factors.get('Revenue', 1):.3f}")

        apply_scaling_to_methodology(methodology_path, factors, methodology_path)

    print(f"\n{'='*100}")
    print(f"CALIBRATION COMPLETE")
    print(f"Best iteration: {best_iteration} with max variance: {best_max_var:.1f}%")
    print(f"Best methodology saved to: sample_data/Rate_Methodology_BEST.csv")
    print(f"{'='*100}")

    return best_variances, best_iteration


if __name__ == '__main__':
    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument('--max-iterations', type=int, default=10)
    parser.add_argument('--tolerance', type=float, default=0.05)
    parser.add_argument('--damping', type=float, default=0.3)
    args = parser.parse_args()

    run_iterative_calibration(
        max_iterations=args.max_iterations,
        tolerance=args.tolerance,
        damping=args.damping
    )
