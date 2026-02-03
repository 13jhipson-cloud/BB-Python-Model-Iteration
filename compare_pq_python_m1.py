"""
Detailed cohort-level comparison between PQ model output and Python model output
for M1 (Oct 2025).
"""
import pandas as pd
import numpy as np

pd.set_option('display.max_rows', 200)
pd.set_option('display.max_columns', 30)
pd.set_option('display.width', 220)
pd.set_option('display.float_format', lambda x: f'{x:,.2f}')

PQ_FILE = '/home/user/BB-Python-Model-Iteration/BackBook Forecast - Baseline Outputs (1) (1).xlsx'
PY_FILE = '/home/user/BB-Python-Model-Iteration/v37_output/Forecast_Transparency_Report.xlsx'

# =============================================================================
# 1. LOAD PQ DATA (BB Segmented sheet, raw with header=None)
# =============================================================================
print("=" * 120)
print("LOADING PQ DATA...")
print("=" * 120)

pq_raw = pd.read_excel(PQ_FILE, sheet_name='BB Segmented', header=None)

# Column 1 = Cohort, Column 2 = Segment, Column 3 = Oct 2025 value
# Segment names have leading/trailing spaces in PQ => strip them

# --- OpeningGBV by Cohort x Segment (rows 358-452) ---
pq_gbv = pq_raw.iloc[358:453, [1, 2, 3]].copy()
pq_gbv.columns = ['Cohort', 'Segment', 'PQ_OpeningGBV']
pq_gbv['Cohort'] = pq_gbv['Cohort'].ffill()  # forward-fill cohort
pq_gbv['Cohort'] = pq_gbv['Cohort'].astype(int)
pq_gbv['Segment'] = pq_gbv['Segment'].str.strip()
pq_gbv = pq_gbv.dropna(subset=['Segment'])
pq_gbv = pq_gbv.reset_index(drop=True)

# --- Coll_Principal by Cohort x Segment (rows 453-547) ---
pq_cp = pq_raw.iloc[453:548, [1, 2, 3]].copy()
pq_cp.columns = ['Cohort', 'Segment', 'PQ_Coll_Principal']
pq_cp['Cohort'] = pq_cp['Cohort'].ffill()
pq_cp['Cohort'] = pq_cp['Cohort'].astype(int)
pq_cp['Segment'] = pq_cp['Segment'].str.strip()
pq_cp = pq_cp.dropna(subset=['Segment'])
pq_cp = pq_cp.reset_index(drop=True)

# --- Coll_Interest by Cohort x Segment (rows 548-642) ---
pq_ci = pq_raw.iloc[548:643, [1, 2, 3]].copy()
pq_ci.columns = ['Cohort', 'Segment', 'PQ_Coll_Interest']
pq_ci['Cohort'] = pq_ci['Cohort'].ffill()
pq_ci['Cohort'] = pq_ci['Cohort'].astype(int)
pq_ci['Segment'] = pq_ci['Segment'].str.strip()
pq_ci = pq_ci.dropna(subset=['Segment'])
pq_ci = pq_ci.reset_index(drop=True)

# --- Segment-level OpeningGBV (rows 6-10) ---
pq_seg_gbv = pq_raw.iloc[6:11, [2, 3]].copy()
pq_seg_gbv.columns = ['Segment', 'PQ_Seg_OpeningGBV']
pq_seg_gbv['Segment'] = pq_seg_gbv['Segment'].str.strip()

# Map PQ segment names to Python segment names
seg_map = {
    'NP': 'NON PRIME',
    'NRP-L': 'NRP-L',
    'NRP-M': 'NRP-M',
    'NRP-S': 'NRP-S',
    'PR': 'PRIME',
}
pq_seg_gbv['Segment'] = pq_seg_gbv['Segment'].map(seg_map)
pq_seg_gbv = pq_seg_gbv.reset_index(drop=True)

# Merge PQ cohort-level data
pq = pq_gbv.merge(pq_cp, on=['Cohort', 'Segment'], how='outer')
pq = pq.merge(pq_ci, on=['Cohort', 'Segment'], how='outer')

# Fill NaN values with 0 for missing combos
pq['PQ_OpeningGBV'] = pq['PQ_OpeningGBV'].fillna(0)
pq['PQ_Coll_Principal'] = pq['PQ_Coll_Principal'].fillna(0)
pq['PQ_Coll_Interest'] = pq['PQ_Coll_Interest'].fillna(0)

print(f"PQ cohort-level data: {len(pq)} rows")
print(f"PQ segments: {sorted(pq['Segment'].unique())}")
print(f"PQ cohorts:  {sorted(pq['Cohort'].unique())}")
print()

# =============================================================================
# 2. LOAD PYTHON DATA (Forecast_Transparency_Report, 10_Details sheet)
# =============================================================================
print("LOADING PYTHON DATA...")
py_all = pd.read_excel(PY_FILE, sheet_name='10_Details')
py = py_all[py_all['ForecastMonth'] == '2025-10-31'].copy()
py = py[['Segment', 'Cohort', 'OpeningGBV', 'Coll_Principal', 'Coll_Interest']].copy()
py.columns = ['Segment', 'Cohort', 'PY_OpeningGBV', 'PY_Coll_Principal', 'PY_Coll_Interest']
py = py.reset_index(drop=True)

print(f"Python M1 data: {len(py)} rows")
print(f"Python segments: {sorted(py['Segment'].unique())}")
print(f"Python cohorts:  {sorted(py['Cohort'].unique())}")
print()

# =============================================================================
# 3. MERGE AND COMPARE
# =============================================================================
print("=" * 120)
print("MERGING PQ AND PYTHON DATA...")
print("=" * 120)

comp = pq.merge(py, on=['Segment', 'Cohort'], how='outer', indicator=True)

# Check merge quality
print(f"\nMerge results:")
print(comp['_merge'].value_counts())
print()

# Show any unmatched rows
if (comp['_merge'] != 'both').any():
    print("UNMATCHED ROWS:")
    print(comp[comp['_merge'] != 'both'][['Segment', 'Cohort', '_merge',
          'PQ_OpeningGBV', 'PY_OpeningGBV']].to_string())
    print()

comp = comp.fillna(0)

# Calculate differences
comp['Diff_OpeningGBV'] = comp['PY_OpeningGBV'] - comp['PQ_OpeningGBV']
comp['Diff_Coll_Principal'] = comp['PY_Coll_Principal'] - comp['PQ_Coll_Principal']
comp['Diff_Coll_Interest'] = comp['PY_Coll_Interest'] - comp['PQ_Coll_Interest']
comp['Diff_Total_Coll'] = comp['Diff_Coll_Principal'] + comp['Diff_Coll_Interest']

# Calculate implied rates (rate = amount / OpeningGBV)
for prefix in ['PQ', 'PY']:
    gbv_col = f'{prefix}_OpeningGBV'
    comp[f'{prefix}_CP_Rate'] = np.where(
        comp[gbv_col] != 0,
        comp[f'{prefix}_Coll_Principal'] / comp[gbv_col],
        0
    )
    comp[f'{prefix}_CI_Rate'] = np.where(
        comp[gbv_col] != 0,
        comp[f'{prefix}_Coll_Interest'] / comp[gbv_col],
        0
    )

comp['Diff_CP_Rate'] = comp['PY_CP_Rate'] - comp['PQ_CP_Rate']
comp['Diff_CI_Rate'] = comp['PY_CI_Rate'] - comp['PQ_CI_Rate']

# Absolute variance for sorting
comp['Abs_Diff_Total_Coll'] = comp['Diff_Total_Coll'].abs()

# =============================================================================
# 4. DISPLAY FULL COMPARISON (sorted by absolute variance)
# =============================================================================
print("=" * 120)
print("FULL COHORT-LEVEL COMPARISON: PQ vs PYTHON (M1 = Oct 2025)")
print("=" * 120)
print()

display_cols = [
    'Segment', 'Cohort',
    'PQ_OpeningGBV', 'PY_OpeningGBV', 'Diff_OpeningGBV',
    'PQ_Coll_Principal', 'PY_Coll_Principal', 'Diff_Coll_Principal',
    'PQ_Coll_Interest', 'PY_Coll_Interest', 'Diff_Coll_Interest',
    'Diff_Total_Coll',
]

comp_sorted = comp.sort_values('Abs_Diff_Total_Coll', ascending=False)

print("--- TOP 20 BIGGEST TOTAL COLLECTION VARIANCES ---")
print()
top20 = comp_sorted.head(20)[display_cols]
print(top20.to_string(index=False))
print()

# =============================================================================
# 5. RATE COMPARISON (TOP 20 by rate difference)
# =============================================================================
print("=" * 120)
print("RATE COMPARISON: IMPLIED RATES = Amount / OpeningGBV (TOP 20 by abs CP+CI rate diff)")
print("=" * 120)
print()

rate_cols = [
    'Segment', 'Cohort', 'PQ_OpeningGBV',
    'PQ_CP_Rate', 'PY_CP_Rate', 'Diff_CP_Rate',
    'PQ_CI_Rate', 'PY_CI_Rate', 'Diff_CI_Rate',
]

comp['Abs_Rate_Diff'] = comp['Diff_CP_Rate'].abs() + comp['Diff_CI_Rate'].abs()
comp_rate_sorted = comp.sort_values('Abs_Rate_Diff', ascending=False)

# Use percentage formatting for rates
fmt_pct = lambda x: f'{x:.6%}'
top20_rate = comp_rate_sorted.head(20)[rate_cols].copy()
for c in ['PQ_CP_Rate', 'PY_CP_Rate', 'Diff_CP_Rate', 'PQ_CI_Rate', 'PY_CI_Rate', 'Diff_CI_Rate']:
    top20_rate[c] = top20_rate[c].apply(fmt_pct)

print(top20_rate.to_string(index=False))
print()

# =============================================================================
# 6. SEGMENT-LEVEL TOTALS
# =============================================================================
print("=" * 120)
print("SEGMENT-LEVEL TOTALS (M1 = Oct 2025)")
print("=" * 120)
print()

seg_totals = comp.groupby('Segment').agg({
    'PQ_OpeningGBV': 'sum',
    'PY_OpeningGBV': 'sum',
    'PQ_Coll_Principal': 'sum',
    'PY_Coll_Principal': 'sum',
    'PQ_Coll_Interest': 'sum',
    'PY_Coll_Interest': 'sum',
}).reset_index()

# Add PQ segment-level GBV for cross-check
seg_totals = seg_totals.merge(pq_seg_gbv, on='Segment', how='left')

seg_totals['Diff_OpeningGBV'] = seg_totals['PY_OpeningGBV'] - seg_totals['PQ_OpeningGBV']
seg_totals['Diff_vs_PQ_Seg_GBV'] = seg_totals['PQ_OpeningGBV'] - seg_totals['PQ_Seg_OpeningGBV'].fillna(0)
seg_totals['Diff_Coll_Principal'] = seg_totals['PY_Coll_Principal'] - seg_totals['PQ_Coll_Principal']
seg_totals['Diff_Coll_Interest'] = seg_totals['PY_Coll_Interest'] - seg_totals['PQ_Coll_Interest']
seg_totals['Diff_Total_Coll'] = seg_totals['Diff_Coll_Principal'] + seg_totals['Diff_Coll_Interest']

# Implied segment rates
for prefix in ['PQ', 'PY']:
    gbv = seg_totals[f'{prefix}_OpeningGBV']
    seg_totals[f'{prefix}_CP_Rate'] = np.where(gbv != 0, seg_totals[f'{prefix}_Coll_Principal'] / gbv, 0)
    seg_totals[f'{prefix}_CI_Rate'] = np.where(gbv != 0, seg_totals[f'{prefix}_Coll_Interest'] / gbv, 0)

seg_totals['Diff_CP_Rate'] = seg_totals['PY_CP_Rate'] - seg_totals['PQ_CP_Rate']
seg_totals['Diff_CI_Rate'] = seg_totals['PY_CI_Rate'] - seg_totals['PQ_CI_Rate']

print("--- OpeningGBV ---")
gbv_cols = ['Segment', 'PQ_Seg_OpeningGBV', 'PQ_OpeningGBV', 'PY_OpeningGBV',
            'Diff_vs_PQ_Seg_GBV', 'Diff_OpeningGBV']
print(seg_totals[gbv_cols].to_string(index=False))
print()

print("--- Coll_Principal ---")
cp_cols = ['Segment', 'PQ_Coll_Principal', 'PY_Coll_Principal', 'Diff_Coll_Principal',
           'PQ_CP_Rate', 'PY_CP_Rate', 'Diff_CP_Rate']
seg_cp = seg_totals[cp_cols].copy()
for c in ['PQ_CP_Rate', 'PY_CP_Rate', 'Diff_CP_Rate']:
    seg_cp[c] = seg_cp[c].apply(fmt_pct)
print(seg_cp.to_string(index=False))
print()

print("--- Coll_Interest ---")
ci_cols = ['Segment', 'PQ_Coll_Interest', 'PY_Coll_Interest', 'Diff_Coll_Interest',
           'PQ_CI_Rate', 'PY_CI_Rate', 'Diff_CI_Rate']
seg_ci = seg_totals[ci_cols].copy()
for c in ['PQ_CI_Rate', 'PY_CI_Rate', 'Diff_CI_Rate']:
    seg_ci[c] = seg_ci[c].apply(fmt_pct)
print(seg_ci.to_string(index=False))
print()

print("--- Total Collections (Coll_Principal + Coll_Interest) ---")
tot_cols = ['Segment', 'PQ_OpeningGBV', 'Diff_Coll_Principal', 'Diff_Coll_Interest', 'Diff_Total_Coll']
print(seg_totals[tot_cols].to_string(index=False))
print()

# Grand totals
print("--- GRAND TOTALS ---")
grand = {
    'PQ_OpeningGBV': comp['PQ_OpeningGBV'].sum(),
    'PY_OpeningGBV': comp['PY_OpeningGBV'].sum(),
    'PQ_Coll_Principal': comp['PQ_Coll_Principal'].sum(),
    'PY_Coll_Principal': comp['PY_Coll_Principal'].sum(),
    'PQ_Coll_Interest': comp['PQ_Coll_Interest'].sum(),
    'PY_Coll_Interest': comp['PY_Coll_Interest'].sum(),
}
grand['Diff_OpeningGBV'] = grand['PY_OpeningGBV'] - grand['PQ_OpeningGBV']
grand['Diff_Coll_Principal'] = grand['PY_Coll_Principal'] - grand['PQ_Coll_Principal']
grand['Diff_Coll_Interest'] = grand['PY_Coll_Interest'] - grand['PQ_Coll_Interest']
grand['Diff_Total_Coll'] = grand['Diff_Coll_Principal'] + grand['Diff_Coll_Interest']

for k, v in grand.items():
    print(f"  {k:25s}: {v:>20,.2f}")
print()

# =============================================================================
# 7. COMPLETE LISTING (all rows, sorted by segment then cohort)
# =============================================================================
print("=" * 120)
print("COMPLETE COHORT-LEVEL LISTING (all rows, sorted by Segment/Cohort)")
print("=" * 120)
print()

full_cols = [
    'Segment', 'Cohort',
    'PQ_OpeningGBV', 'PY_OpeningGBV', 'Diff_OpeningGBV',
    'PQ_Coll_Principal', 'PY_Coll_Principal', 'Diff_Coll_Principal',
    'PQ_Coll_Interest', 'PY_Coll_Interest', 'Diff_Coll_Interest',
]

comp_full = comp.sort_values(['Segment', 'Cohort'])[full_cols]
print(comp_full.to_string(index=False))
print()

print("=" * 120)
print("COMPARISON COMPLETE")
print("=" * 120)
