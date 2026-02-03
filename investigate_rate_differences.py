"""
Investigate rate differences between PQ and Python for top variance cohorts.

Compares:
- Python rates from v37_output/Forecast_Transparency_Report.xlsx (sheet 4_Methodology_Applied)
- PQ implied rates back-calculated from BackBook Forecast - Baseline Outputs (1) (1).xlsx (BB Segmented)

PQ implied rate = collection_amount / OpeningGBV
"""

import pandas as pd
import numpy as np
import warnings
warnings.filterwarnings('ignore')

pd.set_option('display.width', 200)
pd.set_option('display.max_columns', 30)
pd.set_option('display.float_format', lambda x: f'{x:.6f}')

# ============================================================
# 1. LOAD PYTHON RATES
# ============================================================
print("=" * 100)
print("LOADING PYTHON RATES (v37 Forecast Transparency Report, sheet 4_Methodology_Applied)")
print("=" * 100)

py_df = pd.read_excel(
    'v37_output/Forecast_Transparency_Report.xlsx',
    '4_Methodology_Applied'
)

py_rates = py_df[['Segment', 'Cohort', 'MOB',
                   'Coll_Principal_Rate', 'Coll_Principal_Approach',
                   'Coll_Interest_Rate', 'Coll_Interest_Approach']].copy()

print(f"Python rates loaded: {len(py_rates)} rows")
print(f"Segments: {sorted(py_rates['Segment'].unique())}")
print(f"Cohorts:  {sorted(py_rates['Cohort'].unique())}")
print()

# ============================================================
# 2. LOAD PQ DATA (BB Segmented sheet)
# ============================================================
print("=" * 100)
print("LOADING PQ DATA (BackBook Forecast - Baseline Outputs, BB Segmented)")
print("=" * 100)

pq_raw = pd.read_excel(
    'BackBook Forecast - Baseline Outputs (1) (1).xlsx',
    'BB Segmented',
    header=None
)

forecast_dates = [pq_raw.iloc[357, c] for c in range(3, 43)]
print(f"Forecast months: {forecast_dates[0].strftime('%Y-%m')} to {forecast_dates[-1].strftime('%Y-%m')}")

def extract_pq_section(df, start_row, end_row, section_name):
    """Extract a cohort x segment section from the PQ sheet."""
    rows = []
    for idx in range(start_row, end_row):
        cohort = df.iloc[idx, 1]
        segment = df.iloc[idx, 2]
        if pd.isna(cohort) or pd.isna(segment):
            continue
        cohort = int(cohort)
        segment = str(segment).strip()
        values = [df.iloc[idx, c] for c in range(3, 43)]
        rows.append({'Cohort': cohort, 'Segment': segment, 'values': values})
    result = pd.DataFrame(rows)
    print(f"  {section_name}: {len(result)} cohort-segment combinations")
    return result

# OpeningGBV: rows 358..452
pq_gbv = extract_pq_section(pq_raw, 358, 453, "OpeningGBV")
# Coll_Principal: rows 453..547
pq_cp = extract_pq_section(pq_raw, 453, 548, "Coll_Principal")
# Coll_Interest: rows 548..642
pq_ci = extract_pq_section(pq_raw, 548, 643, "Coll_Interest")

print()

# ============================================================
# 3. COMPUTE PQ IMPLIED RATES (M1 = first forecast month)
# ============================================================
print("=" * 100)
print("COMPUTING PQ IMPLIED RATES FOR M1 (first forecast month: Oct 2025)")
print("=" * 100)

def build_rate_table(pq_amount_df, pq_gbv_df, amount_name):
    """Build a table of implied rates = amount / OpeningGBV for each forecast month."""
    records = []
    for _, row in pq_amount_df.iterrows():
        cohort = row['Cohort']
        segment = row['Segment']
        amounts = row['values']

        gbv_match = pq_gbv_df[(pq_gbv_df['Cohort'] == cohort) & (pq_gbv_df['Segment'] == segment)]
        if len(gbv_match) == 0:
            continue
        gbv_vals = gbv_match.iloc[0]['values']

        for i in range(len(amounts)):
            amt = amounts[i]
            gbv = gbv_vals[i]
            if pd.notna(amt) and pd.notna(gbv) and gbv != 0:
                rate = amt / gbv
            else:
                rate = 0.0
            records.append({
                'Cohort': cohort,
                'Segment': segment,
                'ForecastMonth': forecast_dates[i],
                'MonthIndex': i + 1,
                f'PQ_{amount_name}': amt if pd.notna(amt) else 0,
                f'PQ_{amount_name}_GBV': gbv if pd.notna(gbv) else 0,
                f'PQ_{amount_name}_Rate': rate,
            })
    return pd.DataFrame(records)

pq_cp_rates = build_rate_table(pq_cp, pq_gbv, 'Coll_Principal')
pq_ci_rates = build_rate_table(pq_ci, pq_gbv, 'Coll_Interest')

pq_rates = pq_cp_rates.merge(
    pq_ci_rates[['Cohort', 'Segment', 'MonthIndex',
                  'PQ_Coll_Interest', 'PQ_Coll_Interest_GBV', 'PQ_Coll_Interest_Rate']],
    on=['Cohort', 'Segment', 'MonthIndex'],
    how='outer'
)

pq_m1 = pq_rates[pq_rates['MonthIndex'] == 1].copy()
print(f"PQ M1 rates computed: {len(pq_m1)} cohort-segment combinations")
print()

# ============================================================
# 4. MERGE PYTHON AND PQ RATES
# ============================================================
print("=" * 100)
print("MERGING PYTHON AND PQ RATES")
print("=" * 100)

def compute_mob_for_oct2025(cohort):
    year = cohort // 100
    month = cohort % 100
    mob = (2025 - year) * 12 + (10 - month) + 1
    return mob

for c in [201912, 202001, 202101, 202301, 202509]:
    expected = compute_mob_for_oct2025(c)
    actual_rows = py_rates[py_rates['Cohort'] == c]
    if len(actual_rows) > 0:
        actual_mob = actual_rows['MOB'].iloc[0]
        print(f"  Cohort {c}: expected MOB={expected}, actual MOB in Python={actual_mob}")

print()

merged = py_rates.merge(
    pq_m1[['Cohort', 'Segment', 'PQ_Coll_Principal', 'PQ_Coll_Principal_GBV',
            'PQ_Coll_Principal_Rate', 'PQ_Coll_Interest', 'PQ_Coll_Interest_GBV',
            'PQ_Coll_Interest_Rate']],
    on=['Cohort', 'Segment'],
    how='inner'
)

print(f"Merged dataset: {len(merged)} rows")
print()

# ============================================================
# 5. FOCUS ON TOP VARIANCE COHORTS
# ============================================================
target_cohorts = [
    ('NON PRIME', 202509),
    ('NRP-M', 202301),
    ('NRP-M', 202509),
    ('NON PRIME', 202301),
    ('PRIME', 202101),
]

print("=" * 100)
print("DETAILED COMPARISON FOR TOP VARIANCE COHORTS")
print("=" * 100)

for seg, coh in target_cohorts:
    print(f"\n{'~' * 100}")
    print(f"  SEGMENT: {seg}  |  COHORT: {coh}")
    print(f"{'~' * 100}")

    row = merged[(merged['Segment'] == seg) & (merged['Cohort'] == coh)]

    if len(row) == 0:
        print("  *** NOT FOUND IN MERGED DATA ***")
        py_match = py_rates[(py_rates['Segment'] == seg) & (py_rates['Cohort'] == coh)]
        pq_match = pq_m1[(pq_m1['Segment'] == seg) & (pq_m1['Cohort'] == coh)]
        print(f"  Python has {len(py_match)} rows, PQ has {len(pq_match)} rows")
        if len(py_match) > 0:
            print(f"  Python: MOB={py_match['MOB'].iloc[0]}, "
                  f"CP_Rate={py_match['Coll_Principal_Rate'].iloc[0]:.6f} ({py_match['Coll_Principal_Approach'].iloc[0]}), "
                  f"CI_Rate={py_match['Coll_Interest_Rate'].iloc[0]:.6f} ({py_match['Coll_Interest_Approach'].iloc[0]})")
        if len(pq_match) > 0:
            print(f"  PQ: GBV={pq_match['PQ_Coll_Principal_GBV'].iloc[0]:,.2f}, "
                  f"CP={pq_match['PQ_Coll_Principal'].iloc[0]:,.2f}, "
                  f"CI={pq_match['PQ_Coll_Interest'].iloc[0]:,.2f}")
        continue

    r = row.iloc[0]
    mob = r['MOB']

    print(f"  MOB at M1 (Oct 2025): {mob}")
    print()
    print(f"  --- Coll_Principal ---")
    print(f"    Python Approach:  {r['Coll_Principal_Approach']}")
    print(f"    Python Rate:      {r['Coll_Principal_Rate']:>12.6f}")
    print(f"    PQ Implied Rate:  {r['PQ_Coll_Principal_Rate']:>12.6f}")
    diff_cp = r['Coll_Principal_Rate'] - r['PQ_Coll_Principal_Rate']
    print(f"    Difference:       {diff_cp:>12.6f}  ({'Python higher' if diff_cp > 0 else 'PQ higher' if diff_cp < 0 else 'EQUAL'})")
    print(f"    PQ OpeningGBV:    {r['PQ_Coll_Principal_GBV']:>18,.2f}")
    print(f"    PQ Amount:        {r['PQ_Coll_Principal']:>18,.2f}")
    py_amount_cp = r['Coll_Principal_Rate'] * r['PQ_Coll_Principal_GBV']
    print(f"    Python Amount*:   {py_amount_cp:>18,.2f}  (* = Python rate x PQ GBV)")
    print(f"    Amount Diff:      {py_amount_cp - r['PQ_Coll_Principal']:>18,.2f}")
    print()
    print(f"  --- Coll_Interest ---")
    print(f"    Python Approach:  {r['Coll_Interest_Approach']}")
    print(f"    Python Rate:      {r['Coll_Interest_Rate']:>12.6f}")
    print(f"    PQ Implied Rate:  {r['PQ_Coll_Interest_Rate']:>12.6f}")
    diff_ci = r['Coll_Interest_Rate'] - r['PQ_Coll_Interest_Rate']
    print(f"    Difference:       {diff_ci:>12.6f}  ({'Python higher' if diff_ci > 0 else 'PQ higher' if diff_ci < 0 else 'EQUAL'})")
    print(f"    PQ OpeningGBV:    {r['PQ_Coll_Interest_GBV']:>18,.2f}")
    print(f"    PQ Amount:        {r['PQ_Coll_Interest']:>18,.2f}")
    py_amount_ci = r['Coll_Interest_Rate'] * r['PQ_Coll_Interest_GBV']
    print(f"    Python Amount*:   {py_amount_ci:>18,.2f}  (* = Python rate x PQ GBV)")
    print(f"    Amount Diff:      {py_amount_ci - r['PQ_Coll_Interest']:>18,.2f}")

# ============================================================
# 6. INVESTIGATE: PQ ZERO Coll_Interest but Python NON-ZERO
# ============================================================
print("\n")
print("=" * 100)
print("INVESTIGATION: Cohorts where PQ has ZERO Coll_Interest but Python has NON-ZERO rate")
print("=" * 100)

zero_pq_ci = merged[
    (merged['PQ_Coll_Interest'].abs() < 0.01) &
    (merged['Coll_Interest_Rate'].abs() > 1e-6)
].copy()

zero_pq_ci['Py_CI_Abs'] = zero_pq_ci['Coll_Interest_Rate'].abs()
zero_pq_ci = zero_pq_ci.sort_values('Py_CI_Abs', ascending=False)

print(f"\nFound {len(zero_pq_ci)} cohorts where PQ Coll_Interest = 0 but Python rate != 0\n")

if len(zero_pq_ci) > 0:
    print(f"{'Segment':<12} {'Cohort':<8} {'MOB':<5} {'Py_CI_Rate':>12} {'Py_Approach':<16} "
          f"{'PQ_CI_Amt':>14} {'PQ_GBV':>16} {'Implied_Py_Amt':>16}")
    print("-" * 110)
    for _, r in zero_pq_ci.iterrows():
        implied = r['Coll_Interest_Rate'] * r['PQ_Coll_Interest_GBV']
        print(f"{r['Segment']:<12} {r['Cohort']:<8} {r['MOB']:<5.0f} "
              f"{r['Coll_Interest_Rate']:>12.6f} {r['Coll_Interest_Approach']:<16} "
              f"{r['PQ_Coll_Interest']:>14,.2f} {r['PQ_Coll_Interest_GBV']:>16,.2f} "
              f"{implied:>16,.2f}")

    print(f"\n--- Breakdown by Python Approach (for PQ-zero / Python-nonzero Coll_Interest cases) ---")
    approach_counts = zero_pq_ci['Coll_Interest_Approach'].value_counts()
    for approach, count in approach_counts.items():
        subset = zero_pq_ci[zero_pq_ci['Coll_Interest_Approach'] == approach]
        avg_rate = subset['Coll_Interest_Rate'].mean()
        max_rate = subset['Coll_Interest_Rate'].abs().max()
        print(f"  {approach:<20}: {count:>3} cases, avg rate = {avg_rate:.6f}, max |rate| = {max_rate:.6f}")

print()

# ============================================================
# 7. SAME ANALYSIS FOR Coll_Principal (PQ zero, Python non-zero)
# ============================================================
print("=" * 100)
print("INVESTIGATION: Cohorts where PQ has ZERO Coll_Principal but Python has NON-ZERO rate")
print("=" * 100)

zero_pq_cp = merged[
    (merged['PQ_Coll_Principal'].abs() < 0.01) &
    (merged['Coll_Principal_Rate'].abs() > 1e-6)
].copy()

zero_pq_cp['Py_CP_Abs'] = zero_pq_cp['Coll_Principal_Rate'].abs()
zero_pq_cp = zero_pq_cp.sort_values('Py_CP_Abs', ascending=False)

print(f"\nFound {len(zero_pq_cp)} cohorts where PQ Coll_Principal = 0 but Python rate != 0\n")

if len(zero_pq_cp) > 0:
    print(f"{'Segment':<12} {'Cohort':<8} {'MOB':<5} {'Py_CP_Rate':>12} {'Py_Approach':<16} "
          f"{'PQ_CP_Amt':>14} {'PQ_GBV':>16} {'Implied_Py_Amt':>16}")
    print("-" * 110)
    for _, r in zero_pq_cp.iterrows():
        implied = r['Coll_Principal_Rate'] * r['PQ_Coll_Principal_GBV']
        print(f"{r['Segment']:<12} {r['Cohort']:<8} {r['MOB']:<5.0f} "
              f"{r['Coll_Principal_Rate']:>12.6f} {r['Coll_Principal_Approach']:<16} "
              f"{r['PQ_Coll_Principal']:>14,.2f} {r['PQ_Coll_Principal_GBV']:>16,.2f} "
              f"{implied:>16,.2f}")

    print(f"\n--- Breakdown by Python Approach (for PQ-zero / Python-nonzero Coll_Principal cases) ---")
    approach_counts = zero_pq_cp['Coll_Principal_Approach'].value_counts()
    for approach, count in approach_counts.items():
        subset = zero_pq_cp[zero_pq_cp['Coll_Principal_Approach'] == approach]
        avg_rate = subset['Coll_Principal_Rate'].mean()
        max_rate = subset['Coll_Principal_Rate'].abs().max()
        print(f"  {approach:<20}: {count:>3} cases, avg rate = {avg_rate:.6f}, max |rate| = {max_rate:.6f}")

print()

# ============================================================
# 8. FULL COMPARISON TABLE FOR ALL M1 COHORTS
# ============================================================
print("=" * 100)
print("FULL RATE COMPARISON TABLE (sorted by absolute Coll_Interest rate diff)")
print("=" * 100)

merged['CP_RateDiff'] = merged['Coll_Principal_Rate'] - merged['PQ_Coll_Principal_Rate']
merged['CI_RateDiff'] = merged['Coll_Interest_Rate'] - merged['PQ_Coll_Interest_Rate']
merged['CI_RateDiff_Abs'] = merged['CI_RateDiff'].abs()
merged['CP_RateDiff_Abs'] = merged['CP_RateDiff'].abs()

top_ci = merged.sort_values('CI_RateDiff_Abs', ascending=False)

print(f"\n{'Segment':<12} {'Cohort':<8} {'MOB':<5} "
      f"{'Py_CI_Rate':>11} {'PQ_CI_Rate':>11} {'CI_Diff':>11} "
      f"{'Py_CP_Rate':>11} {'PQ_CP_Rate':>11} {'CP_Diff':>11} "
      f"{'CI_Approach':<16} {'CP_Approach':<16}")
print("-" * 145)

for _, r in top_ci.head(30).iterrows():
    print(f"{r['Segment']:<12} {r['Cohort']:<8} {r['MOB']:<5.0f} "
          f"{r['Coll_Interest_Rate']:>11.6f} {r['PQ_Coll_Interest_Rate']:>11.6f} {r['CI_RateDiff']:>11.6f} "
          f"{r['Coll_Principal_Rate']:>11.6f} {r['PQ_Coll_Principal_Rate']:>11.6f} {r['CP_RateDiff']:>11.6f} "
          f"{r['Coll_Interest_Approach']:<16} {r['Coll_Principal_Approach']:<16}")

# ============================================================
# 9. EXTENDED: M1 through M5 for target cohorts
# ============================================================
print("\n")
print("=" * 100)
print("EXTENDED: PQ rates across first 5 forecast months for target cohorts")
print("=" * 100)

for seg, coh in target_cohorts:
    print(f"\n--- {seg} / {coh} ---")

    pq_multi = pq_rates[(pq_rates['Segment'] == seg) & (pq_rates['Cohort'] == coh)]
    pq_multi = pq_multi[pq_multi['MonthIndex'] <= 5].sort_values('MonthIndex')

    if len(pq_multi) == 0:
        print("  No PQ data found")
        continue

    print(f"  {'Month':>6} {'PQ_GBV':>16} {'PQ_CP_Amt':>14} {'PQ_CP_Rate':>12} "
          f"{'PQ_CI_Amt':>14} {'PQ_CI_Rate':>12}")
    for _, r in pq_multi.iterrows():
        mi = int(r['MonthIndex'])
        dt = forecast_dates[mi - 1].strftime('%Y-%m')
        print(f"  M{mi} {dt} {r['PQ_Coll_Principal_GBV']:>16,.2f} "
              f"{r['PQ_Coll_Principal']:>14,.2f} {r['PQ_Coll_Principal_Rate']:>12.6f} "
              f"{r['PQ_Coll_Interest']:>14,.2f} {r['PQ_Coll_Interest_Rate']:>12.6f}")

    py_match = py_rates[(py_rates['Segment'] == seg) & (py_rates['Cohort'] == coh)]
    if len(py_match) > 0:
        pr = py_match.iloc[0]
        print(f"\n  Python M1: CP_Rate={pr['Coll_Principal_Rate']:.6f} ({pr['Coll_Principal_Approach']}), "
              f"CI_Rate={pr['Coll_Interest_Rate']:.6f} ({pr['Coll_Interest_Approach']})")

# ============================================================
# 10. ROOT CAUSE SUMMARY
# ============================================================
print("\n")
print("=" * 100)
print("ROOT CAUSE SUMMARY")
print("=" * 100)

print("""
Key findings from comparing Python vs PQ rates:

1. ZERO vs NON-ZERO ISSUE (Coll_Interest):
   When PQ produces ZERO Coll_Interest for a cohort (e.g., because the cohort has no
   historical interest collection data at a given MOB), the Python model may still produce
   a non-zero rate via CohortAvg or CohortTrend approaches. These approaches extrapolate
   from historical data of the same or similar cohorts, producing phantom interest amounts.

2. APPROACH-SPECIFIC DISCREPANCIES:
   - CohortTrend: Extrapolates trends from historical data, can produce non-zero rates
     even when the specific cohort never had collections at that MOB.
   - CohortAvg: Averages rates across similar cohorts, which can introduce rates where
     the PQ model (which may use cohort-specific logic) shows zero.
   - Manual: Fixed rates that may differ from PQ's interpolation logic.

3. SIGN/MAGNITUDE DIFFERENCES:
   Even when both models produce non-zero rates, the magnitudes can differ significantly
   due to different curve-fitting or averaging methodologies.
""")

n_ci_mismatch = len(merged[merged['CI_RateDiff'].abs() > 0.001])
n_cp_mismatch = len(merged[merged['CP_RateDiff'].abs() > 0.001])
n_ci_zero_issue = len(zero_pq_ci)
n_cp_zero_issue = len(zero_pq_cp)

print(f"STATISTICS:")
print(f"  Total cohort-segments compared:                   {len(merged)}")
print(f"  Coll_Interest rate diffs > 0.1%:                  {n_ci_mismatch}")
print(f"  Coll_Principal rate diffs > 0.1%:                 {n_cp_mismatch}")
print(f"  PQ zero CI / Python non-zero CI:                  {n_ci_zero_issue}")
print(f"  PQ zero CP / Python non-zero CP:                  {n_cp_zero_issue}")
print()
