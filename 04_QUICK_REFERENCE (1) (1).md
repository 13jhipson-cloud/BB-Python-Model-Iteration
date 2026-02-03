# Complete Backbook Forecasting Model - Quick Reference Guide

## Executive Summary

Build a Python module that forecasts loan portfolio performance (collections, GBV, impairment, NBV) for 12-36 months using historical rate curves and impairment assumptions.

---

## Key Formulas at a Glance

### Collections & Interest Revenue
```
Coll_Principal = OpeningGBV × Coll_Principal_Rate
Coll_Interest = OpeningGBV × Coll_Interest_Rate
InterestRevenue = OpeningGBV × InterestRevenue_Rate / 12
WO_DebtSold = OpeningGBV × WO_DebtSold_Rate
WO_Other = OpeningGBV × WO_Other_Rate
```

### Closing GBV
```
ClosingGBV = OpeningGBV 
           + InterestRevenue 
           - Coll_Principal 
           - Coll_Interest 
           - WO_DebtSold 
           - WO_Other
```

### Impairment
```
Total_Provision_Balance = ClosingGBV × Total_Coverage_Ratio
Total_Provision_Movement = Total_Provision_Balance[t] - Total_Provision_Balance[t-1]

If Debt_Sale_WriteOffs > 0:
  Debt_Sale_Provision_Release = Debt_Sale_Coverage_Ratio × Debt_Sale_WriteOffs
  Debt_Sale_Proceeds = Debt_Sale_Proceeds_Rate × Debt_Sale_WriteOffs

Non_DS_Provision_Movement = Total_Provision_Movement + Debt_Sale_Provision_Release
Gross_Impairment_ExcludingDS = Non_DS_Provision_Movement + WO_Other
Debt_Sale_Impact = Debt_Sale_WriteOffs + Debt_Sale_Provision_Release + Debt_Sale_Proceeds
Net_Impairment = Gross_Impairment_ExcludingDS + Debt_Sale_Impact
```

### Closing NBV
```
ClosingNBV = ClosingGBV - Net_Impairment
```

---

## Rate Calculation Approaches

| Approach | Formula | When to Use |
|----------|---------|------------|
| **CohortAvg** | Average(Rate[MOB-N:MOB-1] where MOB > 3) | Stable, recent trends |
| **CohortTrend** | a + b×MOB (linear regression) | Trending data |
| **DonorCohort** | Rate from donor cohort at same MOB | New cohorts |
| **SegMedian** | Median(Rate[all cohorts] at MOB) | Cross-cohort average |
| **Manual** | Fixed rate (Param1) | Overrides, assumptions |
| **Zero** | 0 | No activity |

---

## Rate Caps

| Metric | Min | Max |
|--------|-----|-----|
| Coll_Principal | -0.15 | 0 |
| Coll_Interest | -0.10 | 0 |
| InterestRevenue | 0.10 | 0.50 |
| WO_DebtSold | 0 | 0.12 |
| WO_Other | 0 | 0.01 |
| Total_Coverage_Ratio | 0.05 | 0.50 |
| Debt_Sale_Coverage_Ratio | 0.50 | 1.00 |

---

## Input Files

### Fact_Raw_Full.csv
- **Required columns:** CalendarMonth, Cohort, Segment, MOB, OpeningGBV, [collections], [write-offs], [interest], ClosingGBV_Reported, DaysInMonth, **Provision_Balance**, **Debt_Sale_WriteOffs**, **Debt_Sale_Provision_Release**, **Debt_Sale_Proceeds**
- **Format:** CSV with M/D/YYYY or MM/DD/YYYY dates
- **Size:** ~7,000 rows

### Rate_Methodology.csv
- **Required columns:** Segment, Cohort, Metric, MOB_Start, MOB_End, Approach, Param1, Param2
- **Format:** CSV
- **Size:** ~100 rows

### Debt_Sale_Schedule.csv (Optional)
- **Required columns:** ForecastMonth, Segment, Cohort, Debt_Sale_WriteOffs, Debt_Sale_Coverage_Ratio, Debt_Sale_Proceeds_Rate
- **Format:** CSV
- **Size:** Variable (only debt sale months)

---

## Output Files

| File | Purpose | Sheets |
|------|---------|--------|
| **Forecast_Summary.xlsx** | High-level summary by month & segment | Summary |
| **Forecast_Details.xlsx** | Complete forecast for all cohorts | All_Cohorts |
| **Impairment_Analysis.xlsx** | Impairment detail & coverage trends | Impairment_Detail, Coverage_Ratios |
| **Validation_Report.xlsx** | Validation checks & reconciliation | Reconciliation, Validation_Checks |

---

## Function Call Sequence

```
1. Load Data
   ├─ load_fact_raw()
   ├─ load_rate_methodology()
   └─ load_debt_sale_schedule()

2. Calculate Curves
   ├─ calculate_curves_base()
   ├─ extend_curves()
   ├─ calculate_impairment_actuals()
   └─ calculate_impairment_curves()

3. Generate Seeds
   ├─ generate_seed_curves()
   └─ generate_impairment_seed()

4. Build Lookups
   ├─ build_rate_lookup()
   │  └─ get_methodology()
   │     └─ apply_approach()
   │        └─ [fn_cohort_avg | fn_cohort_trend | fn_donor_cohort | fn_seg_median]
   │     └─ apply_rate_cap()
   └─ build_impairment_lookup()

5. Run Forecast
   └─ run_forecast()
      └─ run_one_step()

6. Generate Outputs
   ├─ generate_summary_output()
   ├─ generate_details_output()
   ├─ generate_impairment_output()
   ├─ generate_validation_output()
   └─ export_to_excel()
```

---

## Methodology Lookup Specificity Scoring

```
Score = (Segment match ? 8 : 0) 
       + (Cohort match ? 4 : 0) 
       + (Metric match ? 2 : 0) 
       + 1 / (1 + MOB_End - MOB_Start)

Example:
  Rule 1: Segment=ALL,    Cohort=ALL,    Metric=Coll_Principal, MOB=0-999  → Score = 2.001
  Rule 2: Segment=NRP-S,  Cohort=ALL,    Metric=Coll_Principal, MOB=0-999  → Score = 10.001 ✓ WINNER
  Rule 3: Segment=NRP-S,  Cohort=202001, Metric=Coll_Principal, MOB=0-999  → Score = 14.001 ✓ WINNER
```

---

## Validation Checks

```
✓ GBV Reconciliation: ClosingGBV = OpeningGBV + InterestRevenue - Collections - WriteOffs
✓ NBV Reconciliation: ClosingNBV = ClosingGBV - Net_Impairment
✓ No NaN or Infinite Values
✓ Forecast Chain Continuity: ClosingGBV[t] = OpeningGBV[t+1]
✓ Coverage Ratios in Range: [0.05, 0.50]
✓ Rates Within Caps
```

---

## Command Line Usage

```bash
python backbook_forecast.py \
    --fact-raw Fact_Raw_Full.csv \
    --methodology Rate_Methodology.csv \
    --debt-sale Debt_Sale_Schedule.csv \
    --output output/ \
    --months 12
```

---

## Common Pitfalls to Avoid

1. **Date Parsing:** Handle both M/D/YYYY and MM/DD/YYYY formats
2. **Cohort Type:** Keep cohorts as strings throughout (not integers)
3. **Division by Zero:** Check OpeningGBV > 0 before dividing
4. **Rounding:** Use consistent decimal precision (2-6 decimals)
5. **Forecast Chain:** Ensure ClosingGBV[t] = OpeningGBV[t+1]
6. **Rate Caps:** Apply caps AFTER rate calculation, BEFORE amount calculation
7. **Impairment:** Track provision balance separately from GBV
8. **Debt Sales:** Only apply debt sale logic in sale months

---

## Data Type Standards

| Field | Type | Example |
|-------|------|---------|
| ForecastMonth | datetime | 2025-10-31 |
| Segment | string | NRP-S |
| Cohort | string | 202001 |
| MOB | int | 69 |
| OpeningGBV | float | 4571.87 |
| Rate | float | -0.150000 |
| Amount | float | -685.78 |
| Coverage Ratio | float | 0.1200 |

---

## Performance Targets

- Load data: < 1 second
- Calculate curves: < 2 seconds
- Build lookups: < 10 seconds
- Run forecast: < 5 seconds
- Generate outputs: < 5 seconds
- **Total: < 30 seconds** for full dataset

---

## Key Metrics to Track

**By Month:**
- Total OpeningGBV
- Total InterestRevenue
- Total Collections (Principal + Interest)
- Total Write-offs
- Total ClosingGBV
- Average Coverage Ratio
- Total Net Impairment
- Total ClosingNBV

**By Segment:**
- Count of active cohorts
- Sum of ClosingGBV
- Average Coverage Ratio
- Total Net Impairment

**By Cohort:**
- MOB progression
- GBV runoff
- Coverage ratio trend
- Impairment charge

---

## Error Handling

```python
try:
    # Load data
    fact_raw = load_fact_raw(fact_raw_path)
except FileNotFoundError:
    print(f"ERROR: File not found: {fact_raw_path}")
    sys.exit(1)
except ValueError as e:
    print(f"ERROR: Invalid data format: {e}")
    sys.exit(1)

# Validate data
if len(fact_raw) == 0:
    print("ERROR: No data loaded")
    sys.exit(1)

# Check for required columns
required_cols = ['CalendarMonth', 'Cohort', 'Segment', 'MOB', 'OpeningGBV']
missing_cols = [col for col in required_cols if col not in fact_raw.columns]
if missing_cols:
    print(f"ERROR: Missing columns: {missing_cols}")
    sys.exit(1)
```

---

## Testing Checklist

- [ ] Data loads correctly
- [ ] Curves calculated without errors
- [ ] Seeds generated (1 per Segment × Cohort)
- [ ] Rate lookup built (no NaN rates)
- [ ] Forecast runs for all months
- [ ] GBV reconciliation passes
- [ ] NBV reconciliation passes
- [ ] No NaN or infinite values
- [ ] Forecast chain continuous
- [ ] Coverage ratios in range
- [ ] Excel files generated
- [ ] All sheets populated
- [ ] Validation report shows all checks passed

---

## Deliverables

- [ ] backbook_forecast.py (single module)
- [ ] requirements.txt
- [ ] README.md
- [ ] 4 Excel output files
- [ ] Unit tests
- [ ] Integration tests
- [ ] Error handling & logging
- [ ] Command line interface

---

## Contact & Support

For questions or clarifications:
1. Refer to 01_SYSTEM_DESIGN.md for architecture
2. Refer to 02_IMPLEMENTATION_GUIDE.md for detailed specs
3. Refer to 03_EXAMPLE_OUTPUTS.md for output formats
4. Check validation checks in 03_EXAMPLE_OUTPUTS.md for troubleshooting
