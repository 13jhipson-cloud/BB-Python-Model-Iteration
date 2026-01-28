#!/usr/bin/env python3
"""
Cohort Diagnostic Tool

This tool helps validate the forecasting logic by showing:
1. Historical actuals for a specific cohort
2. How rates were derived from historical data
3. What forecast approach was used for each metric
4. Step-by-step calculation breakdown

Usage:
    python cohort_diagnostic.py --segment "NRP-S" --cohort "202001"
"""

import argparse
import pandas as pd
import numpy as np
from backbook_forecast import (
    load_fact_raw, load_rate_methodology, calculate_curves_base,
    get_methodology, clean_cohort, safe_divide, Config
)


def run_cohort_diagnostic(fact_raw_path: str, methodology_path: str,
                          segment: str, cohort: str, forecast_months: int = 3):
    """
    Generate a detailed diagnostic report for a specific cohort.
    """

    print("=" * 80)
    print(f"COHORT DIAGNOSTIC REPORT")
    print(f"Segment: {segment} | Cohort: {cohort}")
    print("=" * 80)

    # Load data
    print("\n[1] LOADING DATA...")
    fact_raw = load_fact_raw(fact_raw_path)
    methodology = load_rate_methodology(methodology_path)

    cohort_str = clean_cohort(cohort)

    # Filter to this cohort
    cohort_data = fact_raw[(fact_raw['Segment'] == segment) &
                           (fact_raw['Cohort'] == cohort_str)].copy()

    if len(cohort_data) == 0:
        print(f"ERROR: No data found for Segment={segment}, Cohort={cohort_str}")
        print(f"Available segments: {fact_raw['Segment'].unique().tolist()}")
        print(f"Available cohorts: {fact_raw['Cohort'].unique().tolist()[:10]}...")
        return

    cohort_data = cohort_data.sort_values('MOB').reset_index(drop=True)

    # ==========================================================================
    # SECTION 2: HISTORICAL ACTUALS
    # ==========================================================================
    print("\n" + "=" * 80)
    print("[2] HISTORICAL ACTUALS (Last 6 months of data)")
    print("=" * 80)

    # Show last 6 months of actuals
    last_6 = cohort_data.tail(6).copy()

    # Calculate actual rates
    last_6['Coll_Principal_Rate'] = last_6.apply(
        lambda r: safe_divide(r['Coll_Principal'], r['OpeningGBV']), axis=1)
    last_6['Coll_Interest_Rate'] = last_6.apply(
        lambda r: safe_divide(r['Coll_Interest'], r['OpeningGBV']), axis=1)
    last_6['InterestRevenue_Rate'] = last_6.apply(
        lambda r: safe_divide(r['InterestRevenue'], r['OpeningGBV']) * 12, axis=1)  # Annualized
    last_6['WO_DebtSold_Rate'] = last_6.apply(
        lambda r: safe_divide(r['WO_DebtSold'], r['OpeningGBV']), axis=1)
    last_6['WO_Other_Rate'] = last_6.apply(
        lambda r: safe_divide(r['WO_Other'], r['OpeningGBV']), axis=1)
    last_6['Coverage_Ratio'] = last_6.apply(
        lambda r: safe_divide(r['Provision_Balance'], r['ClosingGBV_Reported']), axis=1)

    # Display actuals
    display_cols = ['CalendarMonth', 'MOB', 'OpeningGBV', 'Coll_Principal', 'Coll_Interest',
                    'InterestRevenue', 'WO_DebtSold', 'WO_Other', 'ClosingGBV_Reported']

    print("\n--- Actual Amounts ---")
    print(last_6[display_cols].to_string(index=False))

    print("\n--- Actual Rates (calculated as Amount / OpeningGBV) ---")
    rate_cols = ['CalendarMonth', 'MOB', 'OpeningGBV', 'Coll_Principal_Rate', 'Coll_Interest_Rate',
                 'InterestRevenue_Rate', 'WO_DebtSold_Rate', 'WO_Other_Rate', 'Coverage_Ratio']
    rate_display = last_6[rate_cols].copy()
    for col in rate_cols[3:]:
        rate_display[col] = rate_display[col].apply(lambda x: f"{x:.6f}")
    print(rate_display.to_string(index=False))

    # ==========================================================================
    # SECTION 3: RATE DERIVATION
    # ==========================================================================
    print("\n" + "=" * 80)
    print("[3] RATE DERIVATION (How forecast rates are calculated)")
    print("=" * 80)

    # Get the last MOB for forecasting
    last_mob = cohort_data['MOB'].max()
    forecast_mob = last_mob + 1

    print(f"\nLast historical MOB: {last_mob}")
    print(f"First forecast MOB: {forecast_mob}")

    # Calculate curves for all cohorts (needed for some approaches)
    curves = calculate_curves_base(fact_raw)

    # Metrics to analyze
    metrics = ['Coll_Principal', 'Coll_Interest', 'InterestRevenue',
               'WO_DebtSold', 'WO_Other', 'Total_Coverage_Ratio']

    print("\n--- Methodology Lookup Results ---")
    print(f"{'Metric':<25} {'Approach':<15} {'Param1':<12} {'Param2':<12}")
    print("-" * 65)

    methodology_results = {}
    for metric in metrics:
        meth = get_methodology(methodology, segment, cohort_str, forecast_mob, metric)
        methodology_results[metric] = meth
        param1 = str(meth['Param1'])[:10] if meth['Param1'] else '-'
        param2 = str(meth['Param2'])[:10] if meth['Param2'] else '-'
        print(f"{metric:<25} {meth['Approach']:<15} {param1:<12} {param2:<12}")

    # ==========================================================================
    # SECTION 4: RATE CALCULATION BREAKDOWN
    # ==========================================================================
    print("\n" + "=" * 80)
    print("[4] RATE CALCULATION BREAKDOWN")
    print("=" * 80)

    # Filter curves for this cohort
    cohort_curves = curves[(curves['Segment'] == segment) &
                           (curves['Cohort'] == cohort_str)].copy()
    cohort_curves = cohort_curves.sort_values('MOB')

    for metric in metrics:
        meth = methodology_results[metric]
        approach = meth['Approach']
        rate_col = f"{metric}_Rate"

        print(f"\n--- {metric} ---")
        print(f"Approach: {approach}")

        if approach == 'CohortAvg':
            # Show the average calculation
            lookback = int(float(meth['Param1'])) if meth['Param1'] and meth['Param1'] != 'None' else 6
            print(f"Lookback periods: {lookback}")

            # Get data for averaging (MOB > 3)
            avg_data = cohort_curves[(cohort_curves['MOB'] > 3) &
                                     (cohort_curves['MOB'] <= last_mob)].tail(lookback)

            if rate_col in avg_data.columns and len(avg_data) > 0:
                rates = avg_data[rate_col].tolist()
                avg_rate = np.mean(rates)
                print(f"Historical rates used (last {len(rates)} MOBs > 3):")
                for i, (mob, rate) in enumerate(zip(avg_data['MOB'].tolist(), rates)):
                    print(f"  MOB {mob}: {rate:.6f}")
                print(f"Average rate: {avg_rate:.6f}")
            else:
                print(f"  No data available for {rate_col}")

        elif approach == 'CohortTrend':
            # Show the trend calculation
            trend_data = cohort_curves[(cohort_curves['MOB'] > 3) &
                                       (cohort_curves['MOB'] < forecast_mob)]

            if rate_col in trend_data.columns and len(trend_data) >= 2:
                x = trend_data['MOB'].values
                y = trend_data[rate_col].values

                # Linear regression
                n = len(x)
                sum_x = np.sum(x)
                sum_y = np.sum(y)
                sum_xy = np.sum(x * y)
                sum_xx = np.sum(x * x)

                b = (n * sum_xy - sum_x * sum_y) / (n * sum_xx - sum_x * sum_x)
                a = (sum_y - b * sum_x) / n

                predicted = a + b * forecast_mob

                print(f"Linear regression: Rate = {a:.6f} + {b:.6f} × MOB")
                print(f"Data points used: {len(x)}")
                print(f"Predicted rate at MOB {forecast_mob}: {predicted:.6f}")
            else:
                print(f"  Insufficient data for trend calculation")

        elif approach == 'Manual':
            print(f"Fixed rate: {meth['Param1']}")

        elif approach == 'Zero':
            print(f"Rate set to: 0.0")

        elif approach == 'SegMedian':
            # Show segment median calculation
            seg_data = curves[(curves['Segment'] == segment) &
                              (curves['MOB'] == forecast_mob)]
            if rate_col in seg_data.columns and len(seg_data) > 0:
                median_rate = seg_data[rate_col].median()
                print(f"Median rate across {len(seg_data)} cohorts at MOB {forecast_mob}: {median_rate:.6f}")
            else:
                print(f"  No data available for segment median")

    # ==========================================================================
    # SECTION 5: FORECAST CALCULATION WALKTHROUGH
    # ==========================================================================
    print("\n" + "=" * 80)
    print("[5] FORECAST CALCULATION WALKTHROUGH (First forecast month)")
    print("=" * 80)

    # Get starting values from last actual month
    last_actual = cohort_data.iloc[-1]
    opening_gbv = last_actual['ClosingGBV_Reported']
    prior_provision = last_actual['Provision_Balance']

    print(f"\nStarting values (from last actual month):")
    print(f"  OpeningGBV: {opening_gbv:,.2f}")
    print(f"  Prior Provision Balance: {prior_provision:,.2f}")
    print(f"  Forecast MOB: {forecast_mob}")

    # Get forecast rates
    print(f"\nForecast rates (after applying methodology):")

    # Calculate each rate based on methodology
    forecast_rates = {}
    for metric in ['Coll_Principal', 'Coll_Interest', 'InterestRevenue',
                   'WO_DebtSold', 'WO_Other']:
        meth = methodology_results[metric]
        approach = meth['Approach']
        rate_col = f"{metric}_Rate"

        if approach == 'CohortAvg':
            lookback = int(float(meth['Param1'])) if meth['Param1'] and meth['Param1'] != 'None' else 6
            avg_data = cohort_curves[(cohort_curves['MOB'] > 3) &
                                     (cohort_curves['MOB'] <= last_mob)].tail(lookback)
            if rate_col in avg_data.columns and len(avg_data) > 0:
                rate = avg_data[rate_col].mean()
            else:
                rate = 0.0
        elif approach == 'Manual':
            rate = float(meth['Param1']) if meth['Param1'] and meth['Param1'] != 'None' else 0.0
        elif approach == 'Zero':
            rate = 0.0
        else:
            rate = 0.0

        # Apply rate caps
        if metric in Config.RATE_CAPS:
            min_cap, max_cap = Config.RATE_CAPS[metric]
            rate = max(min_cap, min(max_cap, rate))

        forecast_rates[metric] = rate
        print(f"  {metric}_Rate: {rate:.6f}")

    # Calculate forecast amounts
    print(f"\nForecast amounts (Rate × OpeningGBV):")

    coll_principal = opening_gbv * forecast_rates['Coll_Principal']
    coll_interest = opening_gbv * forecast_rates['Coll_Interest']
    interest_revenue = opening_gbv * forecast_rates['InterestRevenue'] / 12  # Monthly
    wo_debt_sold = opening_gbv * forecast_rates['WO_DebtSold']
    wo_other = opening_gbv * forecast_rates['WO_Other']

    print(f"  Coll_Principal = {opening_gbv:,.2f} × {forecast_rates['Coll_Principal']:.6f} = {coll_principal:,.2f}")
    print(f"  Coll_Interest = {opening_gbv:,.2f} × {forecast_rates['Coll_Interest']:.6f} = {coll_interest:,.2f}")
    print(f"  InterestRevenue = {opening_gbv:,.2f} × {forecast_rates['InterestRevenue']:.6f} / 12 = {interest_revenue:,.2f}")
    print(f"  WO_DebtSold = {opening_gbv:,.2f} × {forecast_rates['WO_DebtSold']:.6f} = {wo_debt_sold:,.2f}")
    print(f"  WO_Other = {opening_gbv:,.2f} × {forecast_rates['WO_Other']:.6f} = {wo_other:,.2f}")

    # Calculate ClosingGBV
    print(f"\nClosingGBV calculation:")
    print(f"  ClosingGBV = OpeningGBV + InterestRevenue - |Coll_Principal| - |Coll_Interest| - WO_DebtSold - WO_Other")

    closing_gbv = (opening_gbv + interest_revenue - abs(coll_principal) -
                   abs(coll_interest) - wo_debt_sold - wo_other)

    print(f"  ClosingGBV = {opening_gbv:,.2f} + {interest_revenue:,.2f} - {abs(coll_principal):,.2f} - {abs(coll_interest):,.2f} - {wo_debt_sold:,.2f} - {wo_other:,.2f}")
    print(f"  ClosingGBV = {closing_gbv:,.2f}")

    # Calculate impairment
    print(f"\nImpairment calculation:")

    # Get coverage ratio from methodology
    coverage_meth = methodology_results.get('Total_Coverage_Ratio', {})
    if coverage_meth.get('Approach') == 'Manual':
        coverage_ratio = float(coverage_meth['Param1']) if coverage_meth['Param1'] else 0.12
    else:
        # Use average from curves
        if 'Total_Coverage_Ratio' in cohort_curves.columns:
            coverage_ratio = cohort_curves['Total_Coverage_Ratio'].mean()
        else:
            coverage_ratio = 0.12

    provision_balance = closing_gbv * coverage_ratio
    provision_movement = provision_balance - prior_provision
    net_impairment = provision_movement + wo_other

    print(f"  Coverage Ratio: {coverage_ratio:.6f} ({coverage_ratio*100:.2f}%)")
    print(f"  Provision Balance = ClosingGBV × Coverage Ratio = {closing_gbv:,.2f} × {coverage_ratio:.6f} = {provision_balance:,.2f}")
    print(f"  Provision Movement = New Balance - Prior Balance = {provision_balance:,.2f} - {prior_provision:,.2f} = {provision_movement:,.2f}")
    print(f"  Net Impairment = Provision Movement + WO_Other = {provision_movement:,.2f} + {wo_other:,.2f} = {net_impairment:,.2f}")

    # Calculate ClosingNBV
    closing_nbv = closing_gbv - net_impairment
    print(f"\nClosingNBV calculation:")
    print(f"  ClosingNBV = ClosingGBV - Net_Impairment = {closing_gbv:,.2f} - {net_impairment:,.2f} = {closing_nbv:,.2f}")

    # ==========================================================================
    # SECTION 6: COMPARISON SUMMARY
    # ==========================================================================
    print("\n" + "=" * 80)
    print("[6] ACTUALS vs FORECAST COMPARISON")
    print("=" * 80)

    # Get last 3 actuals and show trend
    last_actuals = cohort_data.tail(3)

    print("\n--- Last 3 Actual Months ---")
    print(f"{'Month':<12} {'MOB':<5} {'OpeningGBV':>15} {'ClosingGBV':>15} {'Runoff %':>10}")
    print("-" * 60)
    for _, row in last_actuals.iterrows():
        runoff_pct = (row['OpeningGBV'] - row['ClosingGBV_Reported']) / row['OpeningGBV'] * 100 if row['OpeningGBV'] > 0 else 0
        print(f"{str(row['CalendarMonth'])[:10]:<12} {row['MOB']:<5} {row['OpeningGBV']:>15,.2f} {row['ClosingGBV_Reported']:>15,.2f} {runoff_pct:>10.2f}%")

    print("\n--- First Forecast Month ---")
    forecast_runoff = (opening_gbv - closing_gbv) / opening_gbv * 100 if opening_gbv > 0 else 0
    print(f"{'Forecast':<12} {forecast_mob:<5} {opening_gbv:>15,.2f} {closing_gbv:>15,.2f} {forecast_runoff:>10.2f}%")

    print("\n" + "=" * 80)
    print("END OF DIAGNOSTIC REPORT")
    print("=" * 80)

    return {
        'cohort_data': cohort_data,
        'methodology_results': methodology_results,
        'forecast_rates': forecast_rates,
        'opening_gbv': opening_gbv,
        'closing_gbv': closing_gbv,
        'net_impairment': net_impairment,
        'closing_nbv': closing_nbv
    }


def main():
    parser = argparse.ArgumentParser(description='Cohort Diagnostic Tool')
    parser.add_argument('--fact-raw', '-f', default='Fact_Raw.xlsx',
                        help='Path to Fact_Raw file')
    parser.add_argument('--methodology', '-m', default='sample_data/Rate_Methodology.csv',
                        help='Path to Rate_Methodology file')
    parser.add_argument('--segment', '-s', required=True,
                        help='Segment to analyze (e.g., NRP-S)')
    parser.add_argument('--cohort', '-c', required=True,
                        help='Cohort to analyze (e.g., 202001)')

    args = parser.parse_args()

    run_cohort_diagnostic(
        fact_raw_path=args.fact_raw,
        methodology_path=args.methodology,
        segment=args.segment,
        cohort=args.cohort
    )


if __name__ == '__main__':
    main()
