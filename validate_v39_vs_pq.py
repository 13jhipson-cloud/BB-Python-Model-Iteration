"""
v39 Forecast Output Validation Against PQ Targets
===================================================
Compares Python v39 forecast output (Forecast_Transparency_Report.xlsx)
against the PQ (PowerQuery / BB Output Raw) M1 targets for Coll_Principal
and Coll_Interest, both at segment level and grand total.
Also displays M2-M12 totals for the full 12-month picture.
"""

import pandas as pd
import numpy as np

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------
XLSX = "/home/user/BB-Python-Model-Iteration/v39_output/Forecast_Transparency_Report.xlsx"

# PQ M1 targets by segment (from BB Output Raw)
PQ_M1 = {
    "NON PRIME": {"CP": -6_037_436.68, "CI": -3_225_125.13},
    "NRP-L":     {"CP":   -164_033.47, "CI":    -66_180.50},
    "NRP-M":     {"CP": -2_914_165.62, "CI": -1_328_056.85},
    "NRP-S":     {"CP": -2_108_685.96, "CI":   -804_105.50},
    "PRIME":     {"CP":   -249_927.40, "CI":     -1_640.72},
}

PQ_M1_TOTAL_CP = -11_474_249.13
PQ_M1_TOTAL_CI =  -5_425_108.70
PQ_M1_TOTAL    = -16_899_357.83

SEGMENT_ORDER = ["NON PRIME", "NRP-L", "NRP-M", "NRP-S", "PRIME"]

# ---------------------------------------------------------------------------
# Load data
# ---------------------------------------------------------------------------
print("=" * 100)
print("v39 FORECAST VALIDATION vs PQ TARGETS")
print("=" * 100)

df_summary = pd.read_excel(XLSX, sheet_name="9_Summary")
df_detail  = pd.read_excel(XLSX, sheet_name="10_Details")

# Ensure ForecastMonth is datetime
df_summary["ForecastMonth"] = pd.to_datetime(df_summary["ForecastMonth"])
df_detail["ForecastMonth"]  = pd.to_datetime(df_detail["ForecastMonth"])

months_sorted = sorted(df_summary["ForecastMonth"].unique())

# ---------------------------------------------------------------------------
# Helper: percentage variance
# ---------------------------------------------------------------------------
def pct_var(actual, target):
    if target == 0:
        return np.nan
    return (actual - target) / abs(target) * 100


# ===========================================================================================
# SECTION 1 : M1 SEGMENT-LEVEL COMPARISON  (Summary sheet)
# ===========================================================================================
print("\n")
print("=" * 100)
print("SECTION 1 : M1 (Oct 2025) SEGMENT-LEVEL COMPARISON  --  9_Summary vs PQ")
print("=" * 100)

m1 = df_summary[df_summary["ForecastMonth"] == months_sorted[0]].copy()
m1 = m1.set_index("Segment").reindex(SEGMENT_ORDER)

header = (
    f"{'Segment':<12} | {'Python CP':>16} {'PQ CP':>16} {'Var%':>8} | "
    f"{'Python CI':>16} {'PQ CI':>16} {'Var%':>8} | "
    f"{'Python Tot':>16} {'PQ Tot':>16} {'Var%':>8}"
)
print(header)
print("-" * len(header))

sum_py_cp = 0
sum_py_ci = 0
sum_pq_cp = 0
sum_pq_ci = 0

for seg in SEGMENT_ORDER:
    py_cp = m1.loc[seg, "Coll_Principal"]
    py_ci = m1.loc[seg, "Coll_Interest"]
    pq_cp = PQ_M1[seg]["CP"]
    pq_ci = PQ_M1[seg]["CI"]

    py_tot = py_cp + py_ci
    pq_tot = pq_cp + pq_ci

    sum_py_cp += py_cp
    sum_py_ci += py_ci
    sum_pq_cp += pq_cp
    sum_pq_ci += pq_ci

    print(
        f"{seg:<12} | {py_cp:>16,.2f} {pq_cp:>16,.2f} {pct_var(py_cp, pq_cp):>7.4f}% | "
        f"{py_ci:>16,.2f} {pq_ci:>16,.2f} {pct_var(py_ci, pq_ci):>7.4f}% | "
        f"{py_tot:>16,.2f} {pq_tot:>16,.2f} {pct_var(py_tot, pq_tot):>7.4f}%"
    )

print("-" * len(header))
sum_py_tot = sum_py_cp + sum_py_ci
sum_pq_tot = sum_pq_cp + sum_pq_ci
print(
    f"{'TOTAL':<12} | {sum_py_cp:>16,.2f} {sum_pq_cp:>16,.2f} {pct_var(sum_py_cp, sum_pq_cp):>7.4f}% | "
    f"{sum_py_ci:>16,.2f} {sum_pq_ci:>16,.2f} {pct_var(sum_py_ci, sum_pq_ci):>7.4f}% | "
    f"{sum_py_tot:>16,.2f} {sum_pq_tot:>16,.2f} {pct_var(sum_py_tot, sum_pq_tot):>7.4f}%"
)

# Also compare against the explicit PQ totals provided
print()
print(f"  PQ stated Coll_Principal total : {PQ_M1_TOTAL_CP:>16,.2f}")
print(f"  Python    Coll_Principal total : {sum_py_cp:>16,.2f}   var = {pct_var(sum_py_cp, PQ_M1_TOTAL_CP):+.6f}%")
print(f"  PQ stated Coll_Interest  total : {PQ_M1_TOTAL_CI:>16,.2f}")
print(f"  Python    Coll_Interest  total : {sum_py_ci:>16,.2f}   var = {pct_var(sum_py_ci, PQ_M1_TOTAL_CI):+.6f}%")
print(f"  PQ stated Total Collections    : {PQ_M1_TOTAL:>16,.2f}")
print(f"  Python    Total Collections    : {sum_py_tot:>16,.2f}   var = {pct_var(sum_py_tot, PQ_M1_TOTAL):+.6f}%")

abs_diff = abs(sum_py_tot - PQ_M1_TOTAL)
print()
if abs_diff < 1.00:
    print(f"  >>> MATCH: Python total within $1.00 of PQ target (diff = ${abs_diff:,.2f})")
elif abs_diff < 10.00:
    print(f"  >>> NEAR MATCH: Python total within $10.00 of PQ target (diff = ${abs_diff:,.2f})")
else:
    print(f"  >>> MISMATCH: Python total differs from PQ target by ${abs_diff:,.2f}")


# ===========================================================================================
# SECTION 2 : M1 SEGMENT-LEVEL COMPARISON  (Detail sheet aggregated)
# ===========================================================================================
print("\n")
print("=" * 100)
print("SECTION 2 : M1 (Oct 2025) SEGMENT-LEVEL COMPARISON  --  10_Details aggregated vs PQ")
print("=" * 100)

m1_det = df_detail[df_detail["ForecastMonth"] == months_sorted[0]].copy()
m1_det_seg = m1_det.groupby("Segment")[["Coll_Principal", "Coll_Interest"]].sum()
m1_det_seg = m1_det_seg.reindex(SEGMENT_ORDER)

header2 = (
    f"{'Segment':<12} | {'Detail CP':>16} {'Summary CP':>16} {'Diff':>12} | "
    f"{'Detail CI':>16} {'Summary CI':>16} {'Diff':>12}"
)
print(header2)
print("-" * len(header2))

for seg in SEGMENT_ORDER:
    det_cp = m1_det_seg.loc[seg, "Coll_Principal"]
    det_ci = m1_det_seg.loc[seg, "Coll_Interest"]
    sum_cp = m1.loc[seg, "Coll_Principal"]
    sum_ci = m1.loc[seg, "Coll_Interest"]
    print(
        f"{seg:<12} | {det_cp:>16,.2f} {sum_cp:>16,.2f} {det_cp - sum_cp:>12,.2f} | "
        f"{det_ci:>16,.2f} {sum_ci:>16,.2f} {det_ci - sum_ci:>12,.2f}"
    )

det_total_cp = m1_det_seg["Coll_Principal"].sum()
det_total_ci = m1_det_seg["Coll_Interest"].sum()
print("-" * len(header2))
print(
    f"{'TOTAL':<12} | {det_total_cp:>16,.2f} {sum_py_cp:>16,.2f} {det_total_cp - sum_py_cp:>12,.2f} | "
    f"{det_total_ci:>16,.2f} {sum_py_ci:>16,.2f} {det_total_ci - sum_py_ci:>12,.2f}"
)
print()
print(f"  Detail-level cohort count for M1: {len(m1_det)}")


# ===========================================================================================
# SECTION 3 : M1-M12 MONTHLY TOTALS  (full picture)
# ===========================================================================================
print("\n")
print("=" * 100)
print("SECTION 3 : M1-M12 MONTHLY TOTALS  --  Full 12-Month Forecast Picture")
print("=" * 100)

monthly = (
    df_summary
    .groupby("ForecastMonth")[["Coll_Principal", "Coll_Interest", "InterestRevenue",
                                "WO_DebtSold", "WO_Other", "OpeningGBV", "ClosingGBV",
                                "Net_Impairment"]]
    .sum()
    .sort_index()
)
monthly["Total_Collections"] = monthly["Coll_Principal"] + monthly["Coll_Interest"]
monthly["Month_Label"] = [f"M{i+1}" for i in range(len(monthly))]

header3 = (
    f"{'Mth':<4} {'Period':<12} | {'Coll_Principal':>16} {'Coll_Interest':>16} "
    f"{'Total_Coll':>16} | {'IntRevenue':>16} {'WO_Other':>12} | "
    f"{'OpenGBV':>16} {'CloseGBV':>16} | {'Net_Impairment':>16}"
)
print(header3)
print("-" * len(header3))

cum_cp = 0
cum_ci = 0
for idx, row in monthly.iterrows():
    lbl = row["Month_Label"]
    period = idx.strftime("%Y-%m-%d")
    cum_cp += row["Coll_Principal"]
    cum_ci += row["Coll_Interest"]
    print(
        f"{lbl:<4} {period:<12} | {row['Coll_Principal']:>16,.2f} {row['Coll_Interest']:>16,.2f} "
        f"{row['Total_Collections']:>16,.2f} | {row['InterestRevenue']:>16,.2f} {row['WO_Other']:>12,.2f} | "
        f"{row['OpeningGBV']:>16,.2f} {row['ClosingGBV']:>16,.2f} | {row['Net_Impairment']:>16,.2f}"
    )

print("-" * len(header3))
print(
    f"{'SUM':<4} {'12-Month':<12} | {monthly['Coll_Principal'].sum():>16,.2f} "
    f"{monthly['Coll_Interest'].sum():>16,.2f} "
    f"{monthly['Total_Collections'].sum():>16,.2f} | "
    f"{monthly['InterestRevenue'].sum():>16,.2f} "
    f"{monthly['WO_Other'].sum():>12,.2f} | "
    f"{'':>16} {'':>16} | "
    f"{monthly['Net_Impairment'].sum():>16,.2f}"
)


# ===========================================================================================
# SECTION 4 : M1-M12 BY SEGMENT
# ===========================================================================================
print("\n")
print("=" * 100)
print("SECTION 4 : 12-MONTH TOTALS BY SEGMENT")
print("=" * 100)

seg_totals = (
    df_summary
    .groupby("Segment")[["Coll_Principal", "Coll_Interest", "InterestRevenue",
                          "WO_Other", "Net_Impairment"]]
    .sum()
    .reindex(SEGMENT_ORDER)
)
seg_totals["Total_Collections"] = seg_totals["Coll_Principal"] + seg_totals["Coll_Interest"]

header4 = (
    f"{'Segment':<12} | {'Coll_Principal':>16} {'Coll_Interest':>16} "
    f"{'Total_Coll':>16} | {'IntRevenue':>16} {'Net_Impairment':>16}"
)
print(header4)
print("-" * len(header4))

for seg in SEGMENT_ORDER:
    r = seg_totals.loc[seg]
    print(
        f"{seg:<12} | {r['Coll_Principal']:>16,.2f} {r['Coll_Interest']:>16,.2f} "
        f"{r['Total_Collections']:>16,.2f} | {r['InterestRevenue']:>16,.2f} {r['Net_Impairment']:>16,.2f}"
    )

print("-" * len(header4))
print(
    f"{'TOTAL':<12} | {seg_totals['Coll_Principal'].sum():>16,.2f} "
    f"{seg_totals['Coll_Interest'].sum():>16,.2f} "
    f"{seg_totals['Total_Collections'].sum():>16,.2f} | "
    f"{seg_totals['InterestRevenue'].sum():>16,.2f} "
    f"{seg_totals['Net_Impairment'].sum():>16,.2f}"
)


# ===========================================================================================
# SECTION 5 : FINAL VERDICT
# ===========================================================================================
print("\n")
print("=" * 100)
print("FINAL VERDICT")
print("=" * 100)
print()
print(f"  Target  (PQ M1 Total Collections)  : {PQ_M1_TOTAL:>16,.2f}")
print(f"  Python  (v39 M1 Total Collections)  : {sum_py_tot:>16,.2f}")
print(f"  Absolute difference                 : {abs_diff:>16,.2f}")
print(f"  Relative difference                 : {pct_var(sum_py_tot, PQ_M1_TOTAL):>+16.6f}%")
print()

# Check each segment
all_seg_ok = True
for seg in SEGMENT_ORDER:
    py_cp = m1.loc[seg, "Coll_Principal"]
    py_ci = m1.loc[seg, "Coll_Interest"]
    pq_cp = PQ_M1[seg]["CP"]
    pq_ci = PQ_M1[seg]["CI"]
    cp_diff = abs(py_cp - pq_cp)
    ci_diff = abs(py_ci - pq_ci)
    if cp_diff > 1.0 or ci_diff > 1.0:
        all_seg_ok = False
        print(f"  {seg}: CP diff=${cp_diff:.2f}, CI diff=${ci_diff:.2f}  <-- CHECK")
    else:
        print(f"  {seg}: CP diff=${cp_diff:.2f}, CI diff=${ci_diff:.2f}  OK")

print()
if all_seg_ok and abs_diff < 1.00:
    print("  RESULT: ALL SEGMENTS AND TOTALS MATCH PQ TARGETS (within $1 tolerance)")
elif abs_diff < 10.00:
    print("  RESULT: NEAR MATCH -- minor rounding differences only")
else:
    print("  RESULT: DIFFERENCES DETECTED -- review segment-level variances above")

print()
print("=" * 100)
