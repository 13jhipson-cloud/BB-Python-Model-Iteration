#!/usr/bin/env python3
"""
Compare Python model vs PQ model for M1 (Oct 2025) forecast outputs.
Analyses OpeningGBV, Coll_Principal, Coll_Interest, and implied rates by segment.
"""

import pandas as pd
import numpy as np
import warnings
warnings.filterwarnings('ignore')

PY_FILE = "/home/user/BB-Python-Model-Iteration/v37_output/Forecast_Transparency_Report.xlsx"
PQ_FILE = "/home/user/BB-Python-Model-Iteration/BackBook Forecast - Baseline Outputs (1) (1).xlsx"

SEG_ORDER = ['NON PRIME', 'NRP-L', 'NRP-M', 'NRP-S', 'PRIME']

# ============================================================
# PART 1: RAW SHEET STRUCTURES
# ============================================================
print("=" * 100)
print("PART 1: RAW SHEET STRUCTURES")
print("=" * 100)

# --- Python model: 9_Summary ---
print("\n" + "-" * 80)
print("PYTHON MODEL -- Sheet: 9_Summary")
print("-" * 80)
df_py_summary = pd.read_excel(PY_FILE, sheet_name="9_Summary")
print(f"Shape: {df_py_summary.shape}")
print(f"Columns: {list(df_py_summary.columns)}")
print(df_py_summary.head(10).to_string(index=False))

# --- Python model: 10_Details (M1 only) ---
print("\n" + "-" * 80)
print("PYTHON MODEL -- Sheet: 10_Details (M1 rows only, key columns)")
print("-" * 80)
df_py_detail = pd.read_excel(PY_FILE, sheet_name="10_Details")
df_py_detail_m1 = df_py_detail[df_py_detail['ForecastMonth'] == '2025-10-31'].copy()
key_cols = ['Segment', 'Cohort', 'MOB', 'OpeningGBV',
            'Coll_Principal_Rate', 'Coll_Principal_Approach',
            'Coll_Interest_Rate', 'Coll_Interest_Approach',
            'Coll_Principal', 'Coll_Interest']
print(f"M1 detail rows: {len(df_py_detail_m1)}")
print(df_py_detail_m1[key_cols].head(10).to_string(index=False))

# --- Python model: 4_Methodology_Applied ---
print("\n" + "-" * 80)
print("PYTHON MODEL -- Sheet: 4_Methodology_Applied (first rows)")
print("-" * 80)
df_py_meth = pd.read_excel(PY_FILE, sheet_name="4_Methodology_Applied")
print(f"Shape: {df_py_meth.shape}")
meth_cols = ['Segment', 'Cohort', 'MOB', 'Coll_Principal_Rate', 'Coll_Principal_Approach',
             'Coll_Interest_Rate', 'Coll_Interest_Approach']
print(df_py_meth[meth_cols].head(10).to_string(index=False))

# --- PQ model: Sheet3 (segment level, M1) ---
print("\n" + "-" * 80)
print("PQ MODEL -- Sheet: Sheet3 (segment-level summary rows)")
print("-" * 80)
df_pq_raw = pd.read_excel(PQ_FILE, sheet_name="Sheet3", header=None)
print(f"Shape: {df_pq_raw.shape}")
# Show the structure of the segment-level block (rows 2-47)
for i in range(2, 48):
    row = df_pq_raw.iloc[i]
    vals = [str(v)[:35] for v in row.values[:8]]
    print(f"  Row {i}: {vals}")

# --- PQ model: BB Segmented (key rows) ---
print("\n" + "-" * 80)
print("PQ MODEL -- Sheet: BB Segmented (key metric rows, M1=Oct 2025)")
print("-" * 80)
df_bb_raw = pd.read_excel(PQ_FILE, sheet_name="BB Segmented", header=None)
print(f"Shape: {df_bb_raw.shape}")
# Show row labels and M1 value
for i in [5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 41, 42, 43, 44, 45]:
    row = df_bb_raw.iloc[i]
    vals = [str(v)[:35] for v in row.values[:5]]
    print(f"  Row {i}: {vals}")


# ============================================================
# PART 2: EXTRACT M1 DATA FROM BOTH MODELS
# ============================================================
print("\n\n" + "=" * 100)
print("PART 2: EXTRACT M1 (Oct 2025) DATA FROM BOTH MODELS")
print("=" * 100)

# ---------- PYTHON MODEL ----------
print("\n--- Python Model: M1 from 9_Summary ---")
py_m1 = df_py_summary[df_py_summary['ForecastMonth'] == '2025-10-31'].set_index('Segment')
py_m1 = py_m1.reindex(SEG_ORDER)
print(py_m1[['OpeningGBV', 'Coll_Principal', 'Coll_Interest', 'InterestRevenue']].to_string())

# ---------- PQ MODEL (Sheet3 - has proper Coll_Principal / Coll_Interest split) ----------
print("\n--- PQ Model: M1 from Sheet3 ---")

# Parse the segment-level block from Sheet3
# Structure: rows 3-7 = OpeningGBV, rows 8-12 = Coll_Principal, rows 13-17 = Coll_Interest
# rows 38-42 = InterestRevenue, rows 43-47 = ClosingGBV
# The date columns start at column index 4 (for segment-level block)

pq_data = {}
# Manually extract from the raw data based on row positions
# Row 2 has headers; col 3 = Segment, col 4 = 2025-10-31
metric_map = {
    'OpeningGBV': range(3, 8),
    'Coll_Principal': range(8, 13),
    'Coll_Interest': range(13, 18),
    'InterestRevenue': range(38, 43),
    'ClosingGBV': range(43, 48),
}

for metric, rows in metric_map.items():
    for row_idx in rows:
        seg = str(df_pq_raw.iloc[row_idx, 3]).strip()
        val = df_pq_raw.iloc[row_idx, 4]  # col 4 = Oct 2025
        if seg not in pq_data:
            pq_data[seg] = {}
        pq_data[seg][metric] = float(val) if pd.notna(val) else 0.0

pq_m1 = pd.DataFrame(pq_data).T
pq_m1 = pq_m1.reindex(SEG_ORDER)
print(pq_m1[['OpeningGBV', 'Coll_Principal', 'Coll_Interest', 'InterestRevenue']].to_string())

# Also get PQ data from BB Segmented for cross-check
print("\n--- PQ Model: M1 from BB Segmented (for cross-check) ---")
# Map segment codes: NP->NON PRIME, PR->PRIME, others same
seg_code_map = {'NP': 'NON PRIME', 'NRP-L': 'NRP-L', 'NRP-M': 'NRP-M', 'NRP-S': 'NRP-S', 'PR': 'PRIME'}
bb_data = {}

# Rows 6-10 = OpeningGBV, 11-15 = Collections excl contra, 41-45 = InterestRevenue
for row_idx in range(6, 11):
    seg_code = str(df_bb_raw.iloc[row_idx, 2]).strip()
    seg = seg_code_map.get(seg_code, seg_code)
    val = float(df_bb_raw.iloc[row_idx, 3])
    bb_data.setdefault(seg, {})['OpeningGBV_BB'] = val

for row_idx in range(11, 16):
    seg_code = str(df_bb_raw.iloc[row_idx, 2]).strip()
    seg = seg_code_map.get(seg_code, seg_code)
    val = float(df_bb_raw.iloc[row_idx, 3])
    bb_data[seg]['Collections_ExclContra_BB'] = val

for row_idx in range(41, 46):
    seg_code = str(df_bb_raw.iloc[row_idx, 2]).strip()
    seg = seg_code_map.get(seg_code, seg_code)
    val = float(df_bb_raw.iloc[row_idx, 3])
    bb_data[seg]['InterestRevenue_BB'] = val

bb_m1 = pd.DataFrame(bb_data).T.reindex(SEG_ORDER)
print(bb_m1.to_string())
print("\nNote: BB Segmented 'Collections excl contra' = Coll_Principal + Coll_Interest (combined)")
print("We'll use Sheet3 for the split.")


# ============================================================
# PART 3: SEGMENT-LEVEL COMPARISON TABLE
# ============================================================
print("\n\n" + "=" * 100)
print("PART 3: SEGMENT-LEVEL COMPARISON -- Python vs PQ (M1 = Oct 2025)")
print("=" * 100)

comp = pd.DataFrame(index=SEG_ORDER)

# OpeningGBV
comp['Py_OpeningGBV'] = py_m1['OpeningGBV'].values
comp['PQ_OpeningGBV'] = pq_m1['OpeningGBV'].values
comp['GBV_Diff'] = comp['Py_OpeningGBV'] - comp['PQ_OpeningGBV']
comp['GBV_Diff_%'] = (comp['GBV_Diff'] / comp['PQ_OpeningGBV'] * 100)

# Coll_Principal
comp['Py_CollPrin'] = py_m1['Coll_Principal'].values
comp['PQ_CollPrin'] = pq_m1['Coll_Principal'].values
comp['CollPrin_Diff'] = comp['Py_CollPrin'] - comp['PQ_CollPrin']
comp['CollPrin_Diff_%'] = (comp['CollPrin_Diff'] / comp['PQ_CollPrin'].abs() * 100)

# Coll_Interest
comp['Py_CollInt'] = py_m1['Coll_Interest'].values
comp['PQ_CollInt'] = pq_m1['Coll_Interest'].values
comp['CollInt_Diff'] = comp['Py_CollInt'] - comp['PQ_CollInt']
comp['CollInt_Diff_%'] = (comp['CollInt_Diff'] / comp['PQ_CollInt'].abs() * 100)

# Implied weighted average rates (Amount / OpeningGBV)
comp['Py_CollPrin_Rate'] = comp['Py_CollPrin'] / comp['Py_OpeningGBV']
comp['PQ_CollPrin_Rate'] = comp['PQ_CollPrin'] / comp['PQ_OpeningGBV']
comp['CollPrin_Rate_Diff_bps'] = (comp['Py_CollPrin_Rate'] - comp['PQ_CollPrin_Rate']) * 10000

comp['Py_CollInt_Rate'] = comp['Py_CollInt'] / comp['Py_OpeningGBV']
comp['PQ_CollInt_Rate'] = comp['PQ_CollInt'] / comp['PQ_OpeningGBV']
comp['CollInt_Rate_Diff_bps'] = (comp['Py_CollInt_Rate'] - comp['PQ_CollInt_Rate']) * 10000

# InterestRevenue
comp['Py_IntRev'] = py_m1['InterestRevenue'].values
comp['PQ_IntRev'] = pq_m1['InterestRevenue'].values
comp['IntRev_Diff'] = comp['Py_IntRev'] - comp['PQ_IntRev']
comp['IntRev_Diff_%'] = (comp['IntRev_Diff'] / comp['PQ_IntRev'].abs() * 100)

comp['Py_IntRev_Rate'] = comp['Py_IntRev'] / comp['Py_OpeningGBV']
comp['PQ_IntRev_Rate'] = comp['PQ_IntRev'] / comp['PQ_OpeningGBV']
comp['IntRev_Rate_Diff_bps'] = (comp['Py_IntRev_Rate'] - comp['PQ_IntRev_Rate']) * 10000


# --- Print the comparison in digestible sections ---

print("\n" + "-" * 100)
print("3A. OPENING GBV COMPARISON")
print("-" * 100)
fmt_gbv = comp[['Py_OpeningGBV', 'PQ_OpeningGBV', 'GBV_Diff', 'GBV_Diff_%']].copy()
fmt_gbv.columns = ['Python GBV', 'PQ GBV', 'Difference', 'Diff %']
print(fmt_gbv.to_string(float_format=lambda x: f"{x:,.2f}"))

total_py_gbv = comp['Py_OpeningGBV'].sum()
total_pq_gbv = comp['PQ_OpeningGBV'].sum()
print(f"\n{'TOTAL':<12} {total_py_gbv:>20,.2f} {total_pq_gbv:>20,.2f} {total_py_gbv - total_pq_gbv:>15,.2f} {(total_py_gbv - total_pq_gbv)/total_pq_gbv*100:>10,.4f}")

print("\n" + "-" * 100)
print("3B. COLL_PRINCIPAL COMPARISON")
print("-" * 100)
fmt_cp = comp[['Py_CollPrin', 'PQ_CollPrin', 'CollPrin_Diff', 'CollPrin_Diff_%',
               'Py_CollPrin_Rate', 'PQ_CollPrin_Rate', 'CollPrin_Rate_Diff_bps']].copy()
fmt_cp.columns = ['Python Amt', 'PQ Amt', 'Diff', 'Diff %', 'Py Rate', 'PQ Rate', 'Rate Diff (bps)']
print(fmt_cp.to_string(float_format=lambda x: f"{x:,.4f}" if abs(x) < 1 else f"{x:,.2f}"))

total_py_cp = comp['Py_CollPrin'].sum()
total_pq_cp = comp['PQ_CollPrin'].sum()
print(f"\n{'TOTAL':<12}  Py: {total_py_cp:>15,.2f}  PQ: {total_pq_cp:>15,.2f}  Diff: {total_py_cp-total_pq_cp:>12,.2f}  ({(total_py_cp-total_pq_cp)/abs(total_pq_cp)*100:+.2f}%)")
print(f"             Py wAvg Rate: {total_py_cp/total_py_gbv:.6f}  PQ wAvg Rate: {total_pq_cp/total_pq_gbv:.6f}  Diff: {(total_py_cp/total_py_gbv - total_pq_cp/total_pq_gbv)*10000:+.2f} bps")

print("\n" + "-" * 100)
print("3C. COLL_INTEREST COMPARISON")
print("-" * 100)
fmt_ci = comp[['Py_CollInt', 'PQ_CollInt', 'CollInt_Diff', 'CollInt_Diff_%',
               'Py_CollInt_Rate', 'PQ_CollInt_Rate', 'CollInt_Rate_Diff_bps']].copy()
fmt_ci.columns = ['Python Amt', 'PQ Amt', 'Diff', 'Diff %', 'Py Rate', 'PQ Rate', 'Rate Diff (bps)']
print(fmt_ci.to_string(float_format=lambda x: f"{x:,.4f}" if abs(x) < 1 else f"{x:,.2f}"))

total_py_ci = comp['Py_CollInt'].sum()
total_pq_ci = comp['PQ_CollInt'].sum()
print(f"\n{'TOTAL':<12}  Py: {total_py_ci:>15,.2f}  PQ: {total_pq_ci:>15,.2f}  Diff: {total_py_ci-total_pq_ci:>12,.2f}  ({(total_py_ci-total_pq_ci)/abs(total_pq_ci)*100:+.2f}%)")
print(f"             Py wAvg Rate: {total_py_ci/total_py_gbv:.6f}  PQ wAvg Rate: {total_pq_ci/total_pq_gbv:.6f}  Diff: {(total_py_ci/total_py_gbv - total_pq_ci/total_pq_gbv)*10000:+.2f} bps")

print("\n" + "-" * 100)
print("3D. INTEREST REVENUE COMPARISON")
print("-" * 100)
fmt_ir = comp[['Py_IntRev', 'PQ_IntRev', 'IntRev_Diff', 'IntRev_Diff_%',
               'Py_IntRev_Rate', 'PQ_IntRev_Rate', 'IntRev_Rate_Diff_bps']].copy()
fmt_ir.columns = ['Python Amt', 'PQ Amt', 'Diff', 'Diff %', 'Py Rate', 'PQ Rate', 'Rate Diff (bps)']
print(fmt_ir.to_string(float_format=lambda x: f"{x:,.4f}" if abs(x) < 1 else f"{x:,.2f}"))

total_py_ir = comp['Py_IntRev'].sum()
total_pq_ir = comp['PQ_IntRev'].sum()
print(f"\n{'TOTAL':<12}  Py: {total_py_ir:>15,.2f}  PQ: {total_pq_ir:>15,.2f}  Diff: {total_py_ir-total_pq_ir:>12,.2f}  ({(total_py_ir-total_pq_ir)/abs(total_pq_ir)*100:+.2f}%)")


# ============================================================
# PART 4: VARIANCE DECOMPOSITION (GBV effect vs Rate effect)
# ============================================================
print("\n\n" + "=" * 100)
print("PART 4: VARIANCE DECOMPOSITION -- GBV Effect vs Rate Effect")
print("=" * 100)
print("For each metric: Total Diff = GBV Effect + Rate Effect + Interaction")
print("  GBV Effect  = (Py_GBV - PQ_GBV) * PQ_Rate")
print("  Rate Effect = PQ_GBV * (Py_Rate - PQ_Rate)")
print("  Interaction = (Py_GBV - PQ_GBV) * (Py_Rate - PQ_Rate)")
print()

for metric, py_col, pq_col, rate_py, rate_pq in [
    ('Coll_Principal', 'Py_CollPrin', 'PQ_CollPrin', 'Py_CollPrin_Rate', 'PQ_CollPrin_Rate'),
    ('Coll_Interest', 'Py_CollInt', 'PQ_CollInt', 'Py_CollInt_Rate', 'PQ_CollInt_Rate'),
    ('InterestRevenue', 'Py_IntRev', 'PQ_IntRev', 'Py_IntRev_Rate', 'PQ_IntRev_Rate'),
]:
    print(f"\n  --- {metric} ---")
    print(f"  {'Segment':<12} {'Total Diff':>14} {'GBV Effect':>14} {'Rate Effect':>14} {'Interaction':>14} {'GBV Match?':>12}")
    total_d = total_g = total_r = total_i = 0
    for seg in SEG_ORDER:
        gbv_diff = comp.loc[seg, 'GBV_Diff']
        total_diff = comp.loc[seg, py_col] - comp.loc[seg, pq_col]
        gbv_effect = gbv_diff * comp.loc[seg, rate_pq]
        rate_effect = comp.loc[seg, 'PQ_OpeningGBV'] * (comp.loc[seg, rate_py] - comp.loc[seg, rate_pq])
        interaction = gbv_diff * (comp.loc[seg, rate_py] - comp.loc[seg, rate_pq])
        gbv_match = "YES" if abs(comp.loc[seg, 'GBV_Diff_%']) < 0.01 else "NO"
        print(f"  {seg:<12} {total_diff:>14,.2f} {gbv_effect:>14,.2f} {rate_effect:>14,.2f} {interaction:>14,.2f} {gbv_match:>12}")
        total_d += total_diff
        total_g += gbv_effect
        total_r += rate_effect
        total_i += interaction
    print(f"  {'TOTAL':<12} {total_d:>14,.2f} {total_g:>14,.2f} {total_r:>14,.2f} {total_i:>14,.2f}")


# ============================================================
# PART 5: COHORT-LEVEL DETAIL COMPARISON (Sheet3 vs 10_Details)
# ============================================================
print("\n\n" + "=" * 100)
print("PART 5: COHORT-LEVEL DETAIL COMPARISON -- Python 10_Details vs PQ Sheet3")
print("=" * 100)

# Parse PQ Sheet3 cohort-level data
# Starts at row 50: header row  (Values, Cohort, Segment, dates...)
# Row 50 = header, rows 51+ = data
# Structure: col 2 = Values (metric), col 3 = Cohort, col 4 = Segment, col 5+ = dates

# Find where cohort-level block starts
pq_cohort_start = 50
print(f"\nPQ Sheet3 cohort block header (row {pq_cohort_start}):")
print(f"  {[str(v)[:25] for v in df_pq_raw.iloc[pq_cohort_start].values[:10]]}")

# Parse the cohort-level data
pq_cohort_data = []
current_metric = None
i = pq_cohort_start + 1
while i < len(df_pq_raw):
    row = df_pq_raw.iloc[i]
    metric_val = row[2]
    if pd.notna(metric_val) and str(metric_val).startswith('Sum of '):
        current_metric = str(metric_val).replace('Sum of ', '')
    cohort = row[3]
    segment = row[4]
    oct_val = row[5]  # col 5 = 2025-10-31 for cohort block
    
    if pd.notna(segment) and current_metric:
        pq_cohort_data.append({
            'Metric': current_metric,
            'Cohort': int(cohort) if pd.notna(cohort) else None,
            'Segment': str(segment).strip(),
            'PQ_Value': float(oct_val) if pd.notna(oct_val) else 0.0
        })
    elif pd.notna(cohort) and pd.isna(segment):
        # This row might have a new cohort but segment is in a merged cell
        pass
    
    i += 1

# Forward-fill cohort
pq_cdf = pd.DataFrame(pq_cohort_data)
pq_cdf['Cohort'] = pq_cdf['Cohort'].ffill()
pq_cdf['Cohort'] = pq_cdf['Cohort'].astype(int)

# Pivot to get one row per Segment+Cohort
pq_pivot = pq_cdf.pivot_table(index=['Segment', 'Cohort'], columns='Metric', values='PQ_Value', aggfunc='first').reset_index()

# Python model M1 detail
py_detail = df_py_detail_m1[['Segment', 'Cohort', 'OpeningGBV', 'Coll_Principal', 'Coll_Interest', 'InterestRevenue']].copy()
py_detail['Cohort'] = py_detail['Cohort'].astype(int)
py_detail = py_detail.rename(columns={
    'OpeningGBV': 'Py_OpeningGBV',
    'Coll_Principal': 'Py_Coll_Principal',
    'Coll_Interest': 'Py_Coll_Interest',
    'InterestRevenue': 'Py_InterestRevenue'
})

# Merge
merged = py_detail.merge(pq_pivot, on=['Segment', 'Cohort'], how='outer', suffixes=('', '_pq'))

# Rename PQ columns
if 'OpeningGBV' in merged.columns:
    merged = merged.rename(columns={
        'OpeningGBV': 'PQ_OpeningGBV',
        'Coll_Principal': 'PQ_Coll_Principal',
        'Coll_Interest': 'PQ_Coll_Interest',
        'InterestRevenue': 'PQ_InterestRevenue',
    })

# Fill NaN with 0 for comparison
for c in ['PQ_OpeningGBV', 'PQ_Coll_Principal', 'PQ_Coll_Interest', 'PQ_InterestRevenue',
          'Py_OpeningGBV', 'Py_Coll_Principal', 'Py_Coll_Interest', 'Py_InterestRevenue']:
    if c in merged.columns:
        merged[c] = merged[c].fillna(0)

# Compute diffs
merged['GBV_Diff'] = merged['Py_OpeningGBV'] - merged['PQ_OpeningGBV']
merged['GBV_Diff_%'] = np.where(merged['PQ_OpeningGBV'] != 0,
                                 merged['GBV_Diff'] / merged['PQ_OpeningGBV'] * 100, np.nan)
merged['CollPrin_Diff'] = merged['Py_Coll_Principal'] - merged['PQ_Coll_Principal']
merged['CollInt_Diff'] = merged['Py_Coll_Interest'] - merged['PQ_Coll_Interest']

# Print by segment
for seg in SEG_ORDER:
    seg_data = merged[merged['Segment'] == seg].sort_values('Cohort')
    if len(seg_data) == 0:
        continue
    print(f"\n  --- Segment: {seg} ({len(seg_data)} cohorts) ---")
    print(f"  {'Cohort':>8} {'Py_GBV':>15} {'PQ_GBV':>15} {'GBV_Diff':>12} {'GBV%':>8}  {'Py_CollPrin':>13} {'PQ_CollPrin':>13} {'CP_Diff':>12}  {'Py_CollInt':>12} {'PQ_CollInt':>12} {'CI_Diff':>12}")
    for _, r in seg_data.iterrows():
        gbv_pct = f"{r['GBV_Diff_%']:.2f}" if pd.notna(r['GBV_Diff_%']) else "N/A"
        print(f"  {int(r['Cohort']):>8} {r['Py_OpeningGBV']:>15,.2f} {r['PQ_OpeningGBV']:>15,.2f} {r['GBV_Diff']:>12,.2f} {gbv_pct:>8}  {r['Py_Coll_Principal']:>13,.2f} {r['PQ_Coll_Principal']:>13,.2f} {r['CollPrin_Diff']:>12,.2f}  {r['Py_Coll_Interest']:>12,.2f} {r['PQ_Coll_Interest']:>12,.2f} {r['CollInt_Diff']:>12,.2f}")
    # Subtotals
    py_gbv_s = seg_data['Py_OpeningGBV'].sum()
    pq_gbv_s = seg_data['PQ_OpeningGBV'].sum()
    py_cp_s = seg_data['Py_Coll_Principal'].sum()
    pq_cp_s = seg_data['PQ_Coll_Principal'].sum()
    py_ci_s = seg_data['Py_Coll_Interest'].sum()
    pq_ci_s = seg_data['PQ_Coll_Interest'].sum()
    print(f"  {'SUBTOTAL':>8} {py_gbv_s:>15,.2f} {pq_gbv_s:>15,.2f} {py_gbv_s-pq_gbv_s:>12,.2f} {'':>8}  {py_cp_s:>13,.2f} {pq_cp_s:>13,.2f} {py_cp_s-pq_cp_s:>12,.2f}  {py_ci_s:>12,.2f} {pq_ci_s:>12,.2f} {py_ci_s-pq_ci_s:>12,.2f}")


# ============================================================
# PART 6: TOP COHORT-LEVEL VARIANCES
# ============================================================
print("\n\n" + "=" * 100)
print("PART 6: TOP 15 COHORT-LEVEL VARIANCES BY |Coll_Principal Diff|")
print("=" * 100)
merged['abs_CP_Diff'] = merged['CollPrin_Diff'].abs()
top_cp = merged.nlargest(15, 'abs_CP_Diff')
print(f"\n  {'Segment':<12} {'Cohort':>8} {'Py_GBV':>15} {'PQ_GBV':>15} {'GBV%':>8}  {'Py_CollPrin':>13} {'PQ_CollPrin':>13} {'CP_Diff':>12}  {'Py_Rate':>10} {'PQ_Rate':>10}")
for _, r in top_cp.iterrows():
    py_rate = r['Py_Coll_Principal'] / r['Py_OpeningGBV'] if r['Py_OpeningGBV'] != 0 else 0
    pq_rate = r['PQ_Coll_Principal'] / r['PQ_OpeningGBV'] if r['PQ_OpeningGBV'] != 0 else 0
    gbv_pct = f"{r['GBV_Diff_%']:.2f}" if pd.notna(r['GBV_Diff_%']) else "N/A"
    print(f"  {r['Segment']:<12} {int(r['Cohort']):>8} {r['Py_OpeningGBV']:>15,.2f} {r['PQ_OpeningGBV']:>15,.2f} {gbv_pct:>8}  {r['Py_Coll_Principal']:>13,.2f} {r['PQ_Coll_Principal']:>13,.2f} {r['CollPrin_Diff']:>12,.2f}  {py_rate:>10.6f} {pq_rate:>10.6f}")

print("\n\n" + "=" * 100)
print("PART 6B: TOP 15 COHORT-LEVEL VARIANCES BY |Coll_Interest Diff|")
print("=" * 100)
merged['abs_CI_Diff'] = merged['CollInt_Diff'].abs()
top_ci = merged.nlargest(15, 'abs_CI_Diff')
print(f"\n  {'Segment':<12} {'Cohort':>8} {'Py_GBV':>15} {'PQ_GBV':>15} {'GBV%':>8}  {'Py_CollInt':>13} {'PQ_CollInt':>13} {'CI_Diff':>12}  {'Py_Rate':>10} {'PQ_Rate':>10}")
for _, r in top_ci.iterrows():
    py_rate = r['Py_Coll_Interest'] / r['Py_OpeningGBV'] if r['Py_OpeningGBV'] != 0 else 0
    pq_rate = r['PQ_Coll_Interest'] / r['PQ_OpeningGBV'] if r['PQ_OpeningGBV'] != 0 else 0
    gbv_pct = f"{r['GBV_Diff_%']:.2f}" if pd.notna(r['GBV_Diff_%']) else "N/A"
    print(f"  {r['Segment']:<12} {int(r['Cohort']):>8} {r['Py_OpeningGBV']:>15,.2f} {r['PQ_OpeningGBV']:>15,.2f} {gbv_pct:>8}  {r['Py_Coll_Interest']:>13,.2f} {r['PQ_Coll_Interest']:>13,.2f} {r['CollInt_Diff']:>12,.2f}  {py_rate:>10.6f} {pq_rate:>10.6f}")


# ============================================================
# PART 7: GRAND SUMMARY
# ============================================================
print("\n\n" + "=" * 100)
print("PART 7: GRAND SUMMARY")
print("=" * 100)

print(f"""
  TOTAL PORTFOLIO M1 (Oct 2025):
  
  {'Metric':<22} {'Python Model':>18} {'PQ Model':>18} {'Difference':>15} {'Diff %':>10}
  {'='*85}
  {'OpeningGBV':<22} {total_py_gbv:>18,.2f} {total_pq_gbv:>18,.2f} {total_py_gbv-total_pq_gbv:>15,.2f} {(total_py_gbv-total_pq_gbv)/total_pq_gbv*100:>10.4f}%
  {'Coll_Principal':<22} {total_py_cp:>18,.2f} {total_pq_cp:>18,.2f} {total_py_cp-total_pq_cp:>15,.2f} {(total_py_cp-total_pq_cp)/abs(total_pq_cp)*100:>+10.4f}%
  {'Coll_Interest':<22} {total_py_ci:>18,.2f} {total_pq_ci:>18,.2f} {total_py_ci-total_pq_ci:>15,.2f} {(total_py_ci-total_pq_ci)/abs(total_pq_ci)*100:>+10.4f}%
  {'InterestRevenue':<22} {total_py_ir:>18,.2f} {total_pq_ir:>18,.2f} {total_py_ir-total_pq_ir:>15,.2f} {(total_py_ir-total_pq_ir)/abs(total_pq_ir)*100:>+10.4f}%
  {'Total Collections':<22} {total_py_cp+total_py_ci:>18,.2f} {total_pq_cp+total_pq_ci:>18,.2f} {(total_py_cp+total_py_ci)-(total_pq_cp+total_pq_ci):>15,.2f} {((total_py_cp+total_py_ci)-(total_pq_cp+total_pq_ci))/abs(total_pq_cp+total_pq_ci)*100:>+10.4f}%

  KEY FINDINGS:
""")

# Check GBV match
if abs((total_py_gbv - total_pq_gbv) / total_pq_gbv * 100) < 0.01:
    print("  [1] OpeningGBV MATCHES between models (< 0.01% diff)")
    print("      => Variances in collections are PURELY from rate/methodology differences")
else:
    print(f"  [1] OpeningGBV DIFFERS by {total_py_gbv - total_pq_gbv:,.2f} ({(total_py_gbv - total_pq_gbv) / total_pq_gbv * 100:.4f}%)")
    print("      => Part of collection variance comes from GBV differences")

# Which segments drive variance
print("\n  [2] Segment-level variance drivers for Coll_Principal:")
for seg in SEG_ORDER:
    d = comp.loc[seg, 'CollPrin_Diff']
    pct = comp.loc[seg, 'CollPrin_Diff_%']
    bps = comp.loc[seg, 'CollPrin_Rate_Diff_bps']
    print(f"      {seg:<12}: Amount diff = {d:>12,.2f} ({pct:+.2f}%), Rate diff = {bps:+.1f} bps")

print("\n  [3] Segment-level variance drivers for Coll_Interest:")
for seg in SEG_ORDER:
    d = comp.loc[seg, 'CollInt_Diff']
    pct = comp.loc[seg, 'CollInt_Diff_%']
    bps = comp.loc[seg, 'CollInt_Rate_Diff_bps']
    print(f"      {seg:<12}: Amount diff = {d:>12,.2f} ({pct:+.2f}%), Rate diff = {bps:+.1f} bps")

print("\n" + "=" * 100)
print("ANALYSIS COMPLETE")
print("=" * 100)
