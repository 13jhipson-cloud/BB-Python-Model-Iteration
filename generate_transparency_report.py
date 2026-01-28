#!/usr/bin/env python3
"""
Forecast Transparency Report Generator

Generates comprehensive Excel outputs showing the full audit trail:
1. Raw Actuals Data
2. Historical Rate Curves
3. Extended Forecast Curves
4. Methodology Applied
5. Forecast Rates Used
6. Final Forecast Amounts

All sheets are designed for pivot tables and charts.

Usage:
    python generate_transparency_report.py --fact-raw Fact_Raw.xlsx --methodology sample_data/Rate_Methodology.csv
"""

import argparse
import os
import pandas as pd
import numpy as np
from datetime import datetime
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import ColorScaleRule

from backbook_forecast import (
    load_fact_raw, load_rate_methodology, calculate_curves_base,
    extend_curves, generate_seed_curves, build_rate_lookup,
    build_impairment_lookup, run_forecast, get_methodology,
    clean_cohort, safe_divide, Config,
    calculate_impairment_actuals, calculate_impairment_curves,
    calculate_seasonal_factors, get_seasonal_factor
)


def generate_transparency_report(fact_raw_path: str, methodology_path: str,
                                  output_path: str = 'Forecast_Transparency_Report.xlsx',
                                  forecast_months: int = 12):
    """
    Generate comprehensive Excel report showing full forecast audit trail.
    """
    print("=" * 70)
    print("GENERATING FORECAST TRANSPARENCY REPORT")
    print("=" * 70)

    # ==========================================================================
    # STEP 1: Load all data
    # ==========================================================================
    print("\n[1/8] Loading data...")
    fact_raw = load_fact_raw(fact_raw_path)
    methodology = load_rate_methodology(methodology_path)

    # ==========================================================================
    # STEP 1b: Calculate Seasonal Factors
    # ==========================================================================
    print("[1b/8] Calculating seasonal factors...")
    if Config.ENABLE_SEASONALITY:
        seasonal_factors = calculate_seasonal_factors(fact_raw)
        print("  Seasonal factors calculated by segment and month")
        # Create a DataFrame of seasonal factors for output
        seasonal_rows = []
        for segment, factors in seasonal_factors.items():
            for month, factor in factors.items():
                seasonal_rows.append({
                    'Segment': segment,
                    'Month_Number': month,
                    'Month_Name': ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun',
                                   'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'][month-1],
                    'Seasonal_Factor': round(factor, 4),
                    'Interpretation': f"CR typically {'higher' if factor > 1 else 'lower'} than average by {abs(factor-1)*100:.1f}%"
                })
        seasonal_factors_df = pd.DataFrame(seasonal_rows)
    else:
        print("  Seasonality disabled")
        seasonal_factors_df = pd.DataFrame()

    # ==========================================================================
    # STEP 2: Prepare Actuals Data sheet
    # ==========================================================================
    print("[2/8] Preparing actuals data...")

    actuals_df = fact_raw.copy()

    # Calculate actual rates for each row
    actuals_df['Coll_Principal_Rate'] = actuals_df.apply(
        lambda r: safe_divide(r['Coll_Principal'], r['OpeningGBV']), axis=1)
    actuals_df['Coll_Interest_Rate'] = actuals_df.apply(
        lambda r: safe_divide(r['Coll_Interest'], r['OpeningGBV']), axis=1)
    actuals_df['InterestRevenue_Rate_Annual'] = actuals_df.apply(
        lambda r: safe_divide(r['InterestRevenue'], r['OpeningGBV']) * safe_divide(365, r['DaysInMonth'], 12), axis=1)
    actuals_df['WO_DebtSold_Rate'] = actuals_df.apply(
        lambda r: safe_divide(r['WO_DebtSold'], r['OpeningGBV']), axis=1)
    actuals_df['WO_Other_Rate'] = actuals_df.apply(
        lambda r: safe_divide(r['WO_Other'], r['OpeningGBV']), axis=1)
    actuals_df['Coverage_Ratio'] = actuals_df.apply(
        lambda r: safe_divide(r['Provision_Balance'], r['ClosingGBV_Reported']), axis=1)
    actuals_df['GBV_Runoff_Rate'] = actuals_df.apply(
        lambda r: safe_divide(r['OpeningGBV'] - r['ClosingGBV_Reported'], r['OpeningGBV']), axis=1)

    # Select columns for output
    actuals_cols = [
        'CalendarMonth', 'Segment', 'Cohort', 'MOB',
        'OpeningGBV', 'Coll_Principal', 'Coll_Interest', 'InterestRevenue',
        'WO_DebtSold', 'WO_Other', 'ClosingGBV_Reported', 'Provision_Balance',
        'Coll_Principal_Rate', 'Coll_Interest_Rate', 'InterestRevenue_Rate_Annual',
        'WO_DebtSold_Rate', 'WO_Other_Rate', 'Coverage_Ratio', 'GBV_Runoff_Rate'
    ]
    actuals_output = actuals_df[actuals_cols].copy()
    actuals_output['DataType'] = 'Actual'

    # ==========================================================================
    # STEP 3: Calculate Historical Rate Curves
    # ==========================================================================
    print("[3/8] Calculating historical rate curves...")

    curves_base = calculate_curves_base(fact_raw)

    # Prepare curves output
    curves_cols = [
        'Segment', 'Cohort', 'MOB', 'OpeningGBV', 'ClosingGBV_Reported',
        'Coll_Principal_Rate', 'Coll_Interest_Rate', 'InterestRevenue_Rate',
        'WO_DebtSold_Rate', 'WO_Other_Rate', 'Total_Coverage_Ratio'
    ]

    # Ensure all columns exist
    for col in curves_cols:
        if col not in curves_base.columns:
            curves_base[col] = 0.0

    historical_curves = curves_base[curves_cols].copy()
    historical_curves['CurveType'] = 'Historical'

    # ==========================================================================
    # STEP 4: Generate Extended Curves
    # ==========================================================================
    print("[4/8] Generating extended forecast curves...")

    curves_extended = extend_curves(curves_base, forecast_months)

    # Mark which are extended
    extended_curves = curves_extended[curves_cols].copy()
    max_historical_mob = curves_base.groupby(['Segment', 'Cohort'])['MOB'].max().reset_index()
    max_historical_mob.columns = ['Segment', 'Cohort', 'Max_Historical_MOB']

    extended_curves = extended_curves.merge(max_historical_mob, on=['Segment', 'Cohort'], how='left')
    extended_curves['CurveType'] = extended_curves.apply(
        lambda r: 'Extended' if r['MOB'] > r['Max_Historical_MOB'] else 'Historical', axis=1)
    extended_curves = extended_curves.drop(columns=['Max_Historical_MOB'])

    # ==========================================================================
    # STEP 5: Build Methodology Applied sheet
    # ==========================================================================
    print("[5/8] Building methodology lookup...")

    # Generate seeds
    seed = generate_seed_curves(fact_raw)

    # Build rate lookup
    rate_lookup = build_rate_lookup(seed, curves_extended, methodology, forecast_months)

    # Build impairment lookup
    impairment_actuals = calculate_impairment_actuals(fact_raw)
    impairment_curves = calculate_impairment_curves(impairment_actuals)
    impairment_lookup = build_impairment_lookup(seed, impairment_curves, methodology, forecast_months)

    # Prepare methodology output - show which approach was used
    methodology_output = rate_lookup.copy()

    # Add impairment columns
    if len(impairment_lookup) > 0:
        imp_cols = ['Segment', 'Cohort', 'MOB', 'Total_Coverage_Ratio', 'Total_Coverage_Approach']
        imp_cols = [c for c in imp_cols if c in impairment_lookup.columns]
        methodology_output = methodology_output.merge(
            impairment_lookup[imp_cols],
            on=['Segment', 'Cohort', 'MOB'],
            how='left',
            suffixes=('', '_Imp')
        )

    # ==========================================================================
    # STEP 6: Run Forecast
    # ==========================================================================
    print("[6/8] Running forecast...")

    forecast = run_forecast(seed, rate_lookup, impairment_lookup, forecast_months)

    # Prepare forecast output
    forecast_output = forecast.copy()
    forecast_output['DataType'] = 'Forecast'

    # ==========================================================================
    # STEP 7: Create Combined Actuals + Forecast sheet
    # ==========================================================================
    print("[7/8] Creating combined view...")

    # Build actuals rows manually
    actuals_rows = []
    for _, row in actuals_df.iterrows():
        actuals_rows.append({
            'Month': row['CalendarMonth'],
            'Segment': row['Segment'],
            'Cohort': row['Cohort'],
            'MOB': row['MOB'],
            'DataType': 'Actual',
            'OpeningGBV': row['OpeningGBV'],
            'Coll_Principal': row['Coll_Principal'],
            'Coll_Interest': row['Coll_Interest'],
            'InterestRevenue': row['InterestRevenue'],
            'WO_DebtSold': row['WO_DebtSold'],
            'WO_Other': row['WO_Other'],
            'ClosingGBV': row['ClosingGBV_Reported'],
            'Coll_Principal_Rate': row.get('Coll_Principal_Rate', 0),
            'Coll_Interest_Rate': row.get('Coll_Interest_Rate', 0),
            'InterestRevenue_Rate': row.get('InterestRevenue_Rate_Annual', 0),
            'WO_DebtSold_Rate': row.get('WO_DebtSold_Rate', 0),
            'WO_Other_Rate': row.get('WO_Other_Rate', 0),
            # Coverage ratios
            'Provision_Balance': row.get('Provision_Balance', 0),
            'Total_Coverage_Ratio': row.get('Coverage_Ratio', 0),
        })

    # Build forecast rows manually
    forecast_rows = []
    for _, row in forecast_output.iterrows():
        forecast_rows.append({
            'Month': row['ForecastMonth'],
            'Segment': row['Segment'],
            'Cohort': row['Cohort'],
            'MOB': row['MOB'],
            'DataType': 'Forecast',
            'OpeningGBV': row['OpeningGBV'],
            'Coll_Principal': row['Coll_Principal'],
            'Coll_Interest': row['Coll_Interest'],
            'InterestRevenue': row['InterestRevenue'],
            'WO_DebtSold': row['WO_DebtSold'],
            'WO_Other': row['WO_Other'],
            'ClosingGBV': row['ClosingGBV'],
            'Coll_Principal_Rate': row.get('Coll_Principal_Rate', 0),
            'Coll_Interest_Rate': row.get('Coll_Interest_Rate', 0),
            'InterestRevenue_Rate': row.get('InterestRevenue_Rate', 0),
            'WO_DebtSold_Rate': row.get('WO_DebtSold_Rate', 0),
            'WO_Other_Rate': row.get('WO_Other_Rate', 0),
            # Coverage ratios
            'Provision_Balance': row.get('Total_Provision_Balance', 0),
            'Total_Coverage_Ratio': row.get('Total_Coverage_Ratio', 0),
            'Core_Coverage_Ratio': row.get('Core_Coverage_Ratio', 0),
            'Debt_Sale_Coverage_Ratio': row.get('Debt_Sale_Coverage_Ratio', 0),
        })

    combined_df = pd.DataFrame(actuals_rows + forecast_rows)
    combined_df = combined_df.sort_values(['Segment', 'Cohort', 'Month']).reset_index(drop=True)

    # ==========================================================================
    # WRITE TO EXCEL
    # ==========================================================================
    print("\nWriting Excel file...")

    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        # Sheet 1: README / Guide
        readme_data = {
            'Sheet Name': [
                '1_Actuals_Data',
                '2_Historical_Rates',
                '3_Extended_Curves',
                '4_Methodology_Applied',
                '5_Forecast_Output',
                '6_Combined_View',
                '7_Rate_Methodology_Rules',
                '8_Seasonal_Factors'
            ],
            'Description': [
                'Raw historical data with calculated rates for each month',
                'Aggregated rate curves by Segment × Cohort × MOB (historical only)',
                'Rate curves extended for forecast period (Historical + Extended)',
                'Which forecast approach was used for each Segment × Cohort × MOB × Metric',
                'Final forecast output with all calculated amounts',
                'Actuals + Forecast combined for easy comparison and charting',
                'The methodology rules from Rate_Methodology.csv',
                'Seasonal adjustment factors by Segment and Month (for Coverage Ratio)'
            ],
            'Use For': [
                'Pivot tables showing historical trends, validating raw data',
                'Understanding historical rate patterns by cohort age (MOB)',
                'Seeing how rates are projected forward',
                'Auditing which approach (CohortAvg, Manual, etc.) was used',
                'Final forecast numbers for reporting',
                'Building charts showing Actual vs Forecast over time',
                'Reference for methodology rules',
                'Understanding how monthly seasonality affects CR forecasts'
            ]
        }
        readme_df = pd.DataFrame(readme_data)
        readme_df.to_excel(writer, sheet_name='README', index=False)

        # Sheet 2: Actuals Data
        actuals_output.to_excel(writer, sheet_name='1_Actuals_Data', index=False)

        # Sheet 3: Historical Rate Curves
        historical_curves.to_excel(writer, sheet_name='2_Historical_Rates', index=False)

        # Sheet 4: Extended Curves
        extended_curves.to_excel(writer, sheet_name='3_Extended_Curves', index=False)

        # Sheet 5: Methodology Applied
        methodology_output.to_excel(writer, sheet_name='4_Methodology_Applied', index=False)

        # Sheet 6: Forecast Output
        forecast_output.to_excel(writer, sheet_name='5_Forecast_Output', index=False)

        # Sheet 7: Combined View
        combined_df.to_excel(writer, sheet_name='6_Combined_View', index=False)

        # Sheet 8: Methodology Rules
        methodology.to_excel(writer, sheet_name='7_Rate_Methodology_Rules', index=False)

        # Sheet 9: Seasonal Factors
        if len(seasonal_factors_df) > 0:
            seasonal_factors_df.to_excel(writer, sheet_name='8_Seasonal_Factors', index=False)

    print(f"\n{'=' * 70}")
    print(f"SUCCESS! Report saved to: {output_path}")
    print(f"{'=' * 70}")

    print("\nSheets created:")
    print("  - README: Guide to understanding each sheet")
    print("  - 1_Actuals_Data: Raw data with calculated rates")
    print("  - 2_Historical_Rates: Rate curves from historical data")
    print("  - 3_Extended_Curves: Curves extended for forecast")
    print("  - 4_Methodology_Applied: Which approach used for each metric")
    print("  - 5_Forecast_Output: Final forecast amounts")
    print("  - 6_Combined_View: Actuals + Forecast for charting")
    print("  - 7_Rate_Methodology_Rules: Your methodology rules")
    print("  - 8_Seasonal_Factors: Monthly adjustment factors by segment")

    print("\nSuggested Pivot Tables:")
    print("  1. From '6_Combined_View': Rows=Month, Columns=DataType, Values=Sum of ClosingGBV")
    print("  2. From '2_Historical_Rates': Rows=MOB, Columns=Segment, Values=Avg of Coll_Principal_Rate")
    print("  3. From '4_Methodology_Applied': Filter by Metric, see Approach used")

    return output_path


def main():
    parser = argparse.ArgumentParser(description='Generate Forecast Transparency Report')
    parser.add_argument('--fact-raw', '-f', default='Fact_Raw.xlsx',
                        help='Path to Fact_Raw file')
    parser.add_argument('--methodology', '-m', default='sample_data/Rate_Methodology.csv',
                        help='Path to Rate_Methodology file')
    parser.add_argument('--output', '-o', default='Forecast_Transparency_Report.xlsx',
                        help='Output Excel file path')
    parser.add_argument('--months', '-n', type=int, default=12,
                        help='Forecast horizon in months')

    args = parser.parse_args()

    generate_transparency_report(
        fact_raw_path=args.fact_raw,
        methodology_path=args.methodology,
        output_path=args.output,
        forecast_months=args.months
    )


if __name__ == '__main__':
    main()
