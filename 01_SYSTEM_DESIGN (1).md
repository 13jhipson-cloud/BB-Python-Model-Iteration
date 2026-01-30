# Complete Backbook Forecasting Model - System Design Document

## Executive Summary

This document describes a comprehensive Python-based backbook forecasting model that projects loan portfolio performance over 12-36 months. The model calculates:

1. **Collections & Interest Revenue** - using historical rate curves
2. **Gross Book Value (GBV)** - opening GBV + interest revenue - collections - write-offs
3. **Impairment & Provisions** - coverage ratios applied to GBV, with debt sale adjustments
4. **Net Book Value (NBV)** - closing GBV - net impairment

The model produces multiple Excel outputs for visualization and validation.

---

## 1. Model Architecture

### 1.1 Data Flow

```
Fact_Raw_Full.csv (Actuals)
        ↓
    [Load & Parse]
        ↓
    [Calculate Curves_Base] → Historical rates by Segment × Cohort × MOB
        ↓
    [Extend Curves] → Rates extended to forecast period
        ↓
    [Generate Seed] → Starting balances for forecast
        ↓
    [Build Rate Lookup] → Methodology-driven rate selection
        ↓
    [Forecast Loop] → Month-by-month calculations
        ├─ Collections & Interest Revenue
        ├─ Closing GBV calculation
        ├─ Coverage Ratio application
        ├─ Impairment & Provision calculations
        └─ Closing NBV calculation
        ↓
    [Output Generation] → Multiple Excel workbooks
        ├─ Forecast_Summary.xlsx (key metrics)
        ├─ Forecast_Details.xlsx (all calculations)
        ├─ Impairment_Analysis.xlsx (impairment detail)
        └─ Validation_Report.xlsx (checks & reconciliation)
```

### 1.2 Key Components

| Component | Purpose | Input | Output |
|-----------|---------|-------|--------|
| **Data Loader** | Parse CSV files, handle dates/types | CSV files | DataFrames |
| **Curves Calculator** | Compute historical rates | Fact_Raw | Curves_Base |
| **Curves Extender** | Extend rates to forecast period | Curves_Base | Curves_Extended |
| **Seed Generator** | Create forecast starting point | Fact_Raw | Seed_Curves |
| **Methodology Lookup** | Select rate calculation approach | Rate_Methodology | Approach selection |
| **Rate Functions** | Calculate rates by approach | Curves, Methodology | Rates |
| **Forecast Engine** | Month-by-month calculations | Seed, Rates, Coverage | Forecast rows |
| **Output Generator** | Create Excel workbooks | Forecast data | Excel files |

---

## 2. Detailed Calculations

### 2.1 Collections & Interest Revenue (Monthly)

**Formulas:**
```
Coll_Principal = OpeningGBV × Coll_Principal_Rate
Coll_Interest = OpeningGBV × Coll_Interest_Rate
InterestRevenue = OpeningGBV × InterestRevenue_Rate / 12  # Annualized rate → monthly
NewLoanAmount = OpeningGBV × NewLoanAmount_Rate
WO_DebtSold = OpeningGBV × WO_DebtSold_Rate
WO_Other = OpeningGBV × WO_Other_Rate
ContraSettlements_Principal = OpeningGBV × ContraSettlements_Principal_Rate
ContraSettlements_Interest = OpeningGBV × ContraSettlements_Interest_Rate
```

**Rate Selection:**
- Rates come from Rate_Methodology control table
- Specificity scoring: Segment (+8) > Cohort (+4) > Metric (+2) > MOB range width (tiebreaker)
- Approaches: CohortAvg, CohortTrend, DonorCohort, SegMedian, Manual, Zero
- Rate caps applied per metric (e.g., Coll_Principal: -15% to 0%)

### 2.2 Closing GBV Calculation

**Formula:**
```
ClosingGBV = OpeningGBV 
           + InterestRevenue 
           - Coll_Principal 
           - Coll_Interest 
           - WO_DebtSold 
           - WO_Other
```

**Notes:**
- InterestRevenue is positive (increases GBV)
- Collections and write-offs are negative (reduce GBV)
- Contra settlements are tracked separately (not in GBV formula)
- NewLoanAmount is tracked separately (not in GBV formula)

### 2.3 Impairment & Provision Calculations

#### A) Actuals Calculations (Historical Data)

**Coverage Metrics:**
```
Total_Coverage_Ratio = Total_Provision_Balance / Total_ClosingGBV
Debt_Sale_Coverage_Ratio = Debt_Sale_Provision_Release / Debt_Sale_WriteOffs
Debt_Sale_Proceeds_Rate = Debt_Sale_Proceeds / Debt_Sale_WriteOffs
```

**Provision Movements:**
```
Total_Provision_Movement = Total_Provision_Balance[t] - Total_Provision_Balance[t-1]
Non_DS_Provision_Movement = Total_Provision_Movement + Debt_Sale_Provision_Release
```

**Impairment View:**
```
Gross_Impairment_ExcludingDS = Non_DS_Provision_Movement + WO_Other
Debt_Sale_Impact = Debt_Sale_WriteOffs + Debt_Sale_Provision_Release + Debt_Sale_Proceeds
Net_Impairment = Gross_Impairment_ExcludingDS + Debt_Sale_Impact
```

#### B) Forecast Calculations

**Forecast Inputs (from Rate_Methodology or assumptions):**
- Total_Coverage_Ratio (forecast curve)
- Debt_Sale_WriteOffs (from debt sale schedule; 0 in non-sale months)
- Debt_Sale_Coverage_Ratio (forecast assumption)
- Debt_Sale_Proceeds_Rate (forecast assumption)
- WO_Other (forecast curve)

**Forecast Provision Balance:**
```
Forecast_Total_Provision_Balance[t] = Forecast_Total_Coverage_Ratio[t] × Forecast_ClosingGBV[t]
Forecast_Total_Provision_Movement[t] = Forecast_Total_Provision_Balance[t] - Forecast_Total_Provision_Balance[t-1]
```

**Forecast Debt Sale Lines (sale months only):**
```
Forecast_Debt_Sale_Provision_Release[t] = Forecast_Debt_Sale_Coverage_Ratio[t] × Forecast_Debt_Sale_WriteOffs[t]
Forecast_Debt_Sale_Proceeds[t] = Forecast_Debt_Sale_Proceeds_Rate[t] × Forecast_Debt_Sale_WriteOffs[t]
```

**Back-Solved Core Coverage (month before debt sale):**
```
Implied_DS_Provision = Forecast_Debt_Sale_Coverage_Ratio[t+1] × Forecast_Debt_Sale_WriteOffs[t+1]
Core_Coverage_Ratio[t] = (Forecast_Total_Provision_Balance[t] - Implied_DS_Provision) 
                         / (Forecast_ClosingGBV[t] - Forecast_Debt_Sale_WriteOffs[t+1])
```

**Forecast Impairment Outputs:**
```
Forecast_Non_DS_Provision_Movement = Forecast_Total_Provision_Movement + Forecast_Debt_Sale_Provision_Release
Forecast_Gross_Impairment_ExcludingDS = Forecast_Non_DS_Provision_Movement + Forecast_WO_Other
Forecast_Debt_Sale_Impact = Forecast_Debt_Sale_WriteOffs 
                            + Forecast_Debt_Sale_Provision_Release 
                            + Forecast_Debt_Sale_Proceeds
Forecast_Net_Impairment = Forecast_Gross_Impairment_ExcludingDS + Forecast_Debt_Sale_Impact
```

### 2.4 Closing NBV Calculation

**Formula:**
```
ClosingNBV = ClosingGBV - Net_Impairment
```

**Notes:**
- Net_Impairment is the total impairment charge for the month
- Can be positive (provision increase) or negative (provision release)
- NBV represents the book value after accounting for expected losses

---

## 3. Input Data Requirements

### 3.1 Fact_Raw_Full.csv (Historical Data)

**Required Columns:**
- `CalendarMonth` (string: M/D/YYYY or MM/DD/YYYY)
- `Cohort` (integer: YYYYMM format)
- `Segment` (string: NON PRIME, NRP-S, NRP-M, NRP-L, PRIME)
- `MOB` (integer: months on book)
- `OpeningGBV` (float)
- `NewLoanAmount` (float)
- `Coll_Principal` (float)
- `Coll_Interest` (float)
- `InterestRevenue` (float)
- `WO_DebtSold` (float)
- `WO_Other` (float)
- `ContraSettlements_Principal` (float)
- `ContraSettlements_Interest` (float)
- `ClosingGBV_Reported` (float)
- `DaysInMonth` (integer)
- `Provision_Balance` (float) - **NEW: For impairment calculations**
- `Debt_Sale_WriteOffs` (float) - **NEW: For debt sale tracking**
- `Debt_Sale_Provision_Release` (float) - **NEW: For debt sale provision**
- `Debt_Sale_Proceeds` (float) - **NEW: For debt sale proceeds**

### 3.2 Rate_Methodology.csv (Control Table)

**Required Columns:**
- `Segment` (string: segment name or "ALL")
- `Cohort` (string: cohort YYYYMM or "ALL")
- `Metric` (string: metric name or "ALL")
- `MOB_Start` (integer)
- `MOB_End` (integer)
- `Approach` (string: CohortAvg, CohortTrend, DonorCohort, SegMedian, Manual, Zero)
- `Param1` (string: optional parameter)
- `Param2` (string: optional parameter)

**Example Rows:**
```
Segment,Cohort,Metric,MOB_Start,MOB_End,Approach,Param1,Param2
ALL,ALL,Coll_Principal,0,999,CohortAvg,6,
NON PRIME,202001,Coll_Principal,0,999,Manual,-0.05,
NRP-S,202001,Coll_Principal,16,40,DonorCohort,201901,
ALL,ALL,Total_Coverage_Ratio,0,999,CohortAvg,6,
ALL,202504,Debt_Sale_Coverage_Ratio,0,999,Manual,0.85,
```

### 3.3 Debt_Sale_Schedule.csv (Optional - for debt sale forecasts)

**Columns:**
- `ForecastMonth` (string: M/D/YYYY)
- `Segment` (string)
- `Cohort` (string)
- `Debt_Sale_WriteOffs` (float)
- `Debt_Sale_Coverage_Ratio` (float)
- `Debt_Sale_Proceeds_Rate` (float)

---

## 4. Output Structure

### 4.1 Forecast_Summary.xlsx

**Sheet: Summary**
- Key metrics by ForecastMonth and Segment
- Columns: ForecastMonth, Segment, OpeningGBV, InterestRevenue, Coll_Principal, Coll_Interest, ClosingGBV, Total_Coverage_Ratio, Net_Impairment, ClosingNBV
- Useful for high-level review

### 4.2 Forecast_Details.xlsx

**Sheet: All_Cohorts**
- Complete forecast for all Segment × Cohort combinations
- Columns: ForecastMonth, Segment, Cohort, MOB, OpeningGBV, [all rate columns], [all amount columns], ClosingGBV, Total_Coverage_Ratio, Total_Provision_Balance, Net_Impairment, ClosingNBV

### 4.3 Impairment_Analysis.xlsx

**Sheet: Impairment_Detail**
- Impairment-specific calculations
- Columns: ForecastMonth, Segment, Cohort, MOB, ClosingGBV, Total_Coverage_Ratio, Total_Provision_Balance, Total_Provision_Movement, Debt_Sale_WriteOffs, Debt_Sale_Coverage_Ratio, Debt_Sale_Provision_Release, Debt_Sale_Proceeds, Non_DS_Provision_Movement, Gross_Impairment_ExcludingDS, Debt_Sale_Impact, Net_Impairment

**Sheet: Coverage_Ratios**
- Coverage ratio trends by Segment and Cohort
- Useful for validating coverage assumptions

### 4.4 Validation_Report.xlsx

**Sheet: Reconciliation**
- Month-by-month reconciliation checks
- Columns: ForecastMonth, Segment, Cohort, OpeningGBV, +InterestRevenue, -Coll_Principal, -Coll_Interest, -WO_DebtSold, -WO_Other, =ClosingGBV_Calculated, ClosingGBV_Forecast, Variance

**Sheet: Validation_Checks**
- Data quality checks
- Missing values, outliers, formula errors

---

## 5. Rate Calculation Approaches

### 5.1 CohortAvg
**Description:** Average of last N MOBs (post-MOB 3)
**Parameters:** Param1 = lookback periods (default 6)
**Formula:** `Rate = Average(Rate[MOB-N:MOB-1] where MOB > 3)`

### 5.2 CohortTrend
**Description:** Linear regression extrapolation on post-MOB 3 data
**Parameters:** None
**Formula:** `Rate[MOB] = a + b × MOB` (where a, b from linear regression)

### 5.3 DonorCohort
**Description:** Copy rate from another cohort at same MOB
**Parameters:** Param1 = donor cohort YYYYMM
**Formula:** `Rate = Rate_from_donor_cohort[MOB]`

### 5.4 SegMedian
**Description:** Median rate across all cohorts in segment at MOB
**Parameters:** None
**Formula:** `Rate = Median(Rate[all cohorts in segment] at MOB)`

### 5.5 Manual
**Description:** Fixed rate override
**Parameters:** Param1 = rate value
**Formula:** `Rate = Param1` (bypasses rate caps)

### 5.6 Zero
**Description:** Force rate to zero
**Parameters:** None
**Formula:** `Rate = 0`

---

## 6. Rate Caps

Rates are capped to prevent unrealistic forecasts:

| Metric | Min | Max | Notes |
|--------|-----|-----|-------|
| Coll_Principal | -0.15 | 0 | Collections reduce GBV |
| Coll_Interest | -0.10 | 0 | Interest collections reduce GBV |
| InterestRevenue | 0.10 | 0.50 | Annualized rate |
| WO_DebtSold | 0 | 0.12 | Write-offs reduce GBV |
| WO_Other | 0 | 0.01 | Other write-offs reduce GBV |
| ContraSettlements_Principal | -0.06 | 0 | Contra settlements |
| ContraSettlements_Interest | -0.005 | 0 | Contra settlements |
| NewLoanAmount | 0 | 1.0 | New originations |
| Total_Coverage_Ratio | 0.05 | 0.50 | Provision as % of GBV |
| Debt_Sale_Coverage_Ratio | 0.50 | 1.00 | Provision on debt sales |
| Debt_Sale_Proceeds_Rate | 0.30 | 1.00 | Proceeds as % of write-offs |

**Note:** Manual overrides bypass rate caps.

---

## 7. Methodology Lookup Specificity Scoring

When multiple rules match a (Segment, Cohort, Metric, MOB), the most specific rule wins:

**Scoring:**
- Exact Segment match: +8 points
- Exact Cohort match: +4 points
- Exact Metric match: +2 points
- Narrower MOB range: +1/(1 + MOB_End - MOB_Start) points (tiebreaker)

**Example:**
```
Rule 1: Segment=ALL,    Cohort=ALL,    Metric=Coll_Principal, MOB=0-999  → Score = 0 + 0 + 2 + 0.001 = 2.001
Rule 2: Segment=NRP-S,  Cohort=ALL,    Metric=Coll_Principal, MOB=0-999  → Score = 8 + 0 + 2 + 0.001 = 10.001 ✓ WINNER
Rule 3: Segment=NRP-S,  Cohort=202001, Metric=Coll_Principal, MOB=0-999  → Score = 8 + 4 + 2 + 0.001 = 14.001 ✓ WINNER
```

---

## 8. Forecast Loop Logic

```python
for month in range(max_months):
    for each (Segment, Cohort) in seed:
        # 1. Get opening balance from prior month (or seed if month 0)
        opening_gbv = seed[segment, cohort].bom
        
        # 2. Look up rates for all metrics
        for metric in [Coll_Principal, Coll_Interest, ...]:
            methodology = get_methodology(segment, cohort, mob, metric)
            rate = apply_approach(methodology, curves)
            rate = apply_rate_cap(rate, metric)
        
        # 3. Calculate amounts
        coll_principal = opening_gbv * coll_principal_rate
        coll_interest = opening_gbv * coll_interest_rate
        interest_revenue = opening_gbv * interest_revenue_rate / 12
        # ... other amounts
        
        # 4. Calculate closing GBV
        closing_gbv = opening_gbv + interest_revenue - coll_principal - coll_interest - wo_debtsold - wo_other
        
        # 5. Calculate impairment
        total_provision_balance = closing_gbv * total_coverage_ratio
        total_provision_movement = total_provision_balance - prior_provision_balance
        # ... debt sale calculations
        net_impairment = gross_impairment_excl_ds + debt_sale_impact
        
        # 6. Calculate closing NBV
        closing_nbv = closing_gbv - net_impairment
        
        # 7. Prepare next seed
        next_seed[segment, cohort].bom = closing_gbv
        next_seed[segment, cohort].mob = mob + 1
```

---

## 9. Data Type & Format Standards

| Field | Type | Format | Example |
|-------|------|--------|---------|
| ForecastMonth | date | YYYY-MM-DD | 2025-10-31 |
| Segment | string | Title case | NRP-S |
| Cohort | string | YYYYMM | 202001 |
| MOB | integer | Integer | 69 |
| OpeningGBV | float | 2 decimals | 4571.87 |
| Rate | float | 6 decimals | -0.150000 |
| Amount | float | 2 decimals | -685.78 |
| Coverage Ratio | float | 4 decimals | 0.1250 |
| Approach | string | CamelCase | CohortAvg |

---

## 10. Error Handling & Validation

### 10.1 Data Validation
- Check for missing required columns
- Validate date formats (convert M/D/YYYY and MM/DD/YYYY)
- Ensure numeric columns are numeric
- Check for negative GBV (flag as warning)
- Check for coverage ratios outside 0-1 range

### 10.2 Calculation Validation
- Verify ClosingGBV = OpeningGBV + InterestRevenue - Collections - WriteOffs
- Verify ClosingNBV = ClosingGBV - Net_Impairment
- Check for NaN or infinite values
- Verify forecast chain continuity (ClosingGBV[t] = OpeningGBV[t+1])

### 10.3 Output Validation
- Reconciliation checks in Excel output
- Summary statistics by Segment and Cohort
- Comparison of forecast to actuals (if available)

---

## 11. Configuration Parameters

```python
class Config:
    MAX_MONTHS = 12  # Forecast horizon
    LOOKBACK_PERIODS = 6  # Default for CohortAvg
    MOB_THRESHOLD = 3  # Minimum MOB for rate calculation
    RATE_CAPS = {...}  # Per-metric caps
    SEGMENTS = ['NON PRIME', 'NRP-S', 'NRP-M', 'NRP-L', 'PRIME']
    METRICS = [
        'Coll_Principal', 'Coll_Interest', 'InterestRevenue',
        'WO_DebtSold', 'WO_Other', 'ContraSettlements_Principal',
        'ContraSettlements_Interest', 'NewLoanAmount',
        'Total_Coverage_Ratio', 'Debt_Sale_Coverage_Ratio',
        'Debt_Sale_Proceeds_Rate', 'WO_Other'
    ]
```

---

## 12. Next Steps

1. **Data Preparation:** Ensure Fact_Raw_Full.csv includes Provision_Balance and debt sale columns
2. **Methodology Setup:** Define Rate_Methodology rules for all metrics including coverage ratios
3. **Debt Sale Schedule:** Create schedule for debt sale months and assumptions
4. **Validation Rules:** Define acceptable ranges for key metrics
5. **Excel Templates:** Create output workbook templates for consistency
