# Complete Backbook Forecasting Model - Example Outputs & Validation

## Overview

This document provides examples of expected output formats and validation checks for the complete backbook forecasting model.

---

## Part 1: Forecast_Summary.xlsx Output

### Sheet: Summary

This sheet provides high-level summary by ForecastMonth and Segment.

**Example Data:**

| ForecastMonth | Segment | OpeningGBV | InterestRevenue | Coll_Principal | Coll_Interest | WO_DebtSold | WO_Other | ClosingGBV | Total_Coverage_Ratio | Net_Impairment | ClosingNBV |
|---|---|---|---|---|---|---|---|---|---|---|---|
| 2025-10-31 | NON PRIME | 86,444.65 | 720.37 | -4,322.23 | 0.00 | 0.00 | 0.00 | 82,842.79 | 0.1250 | 10,355.35 | 72,487.44 |
| 2025-10-31 | NRP-S | 4,571.87 | 38.10 | -685.78 | -33.71 | 0.00 | 0.00 | 3,890.48 | 0.1200 | 466.86 | 3,423.62 |
| 2025-10-31 | NRP-M | 0.00 | 0.00 | 0.00 | 0.00 | 0.00 | 0.00 | 0.00 | 0.0000 | 0.00 | 0.00 |
| 2025-11-30 | NON PRIME | 82,842.79 | 690.36 | -4,142.14 | 0.00 | 0.00 | 0.00 | 79,391.01 | 0.1250 | 9,923.88 | 69,467.13 |
| 2025-11-30 | NRP-S | 3,890.48 | 32.42 | -583.57 | -26.37 | 0.00 | 0.00 | 3,312.96 | 0.1200 | 397.56 | 2,915.40 |

**Key Metrics:**
- **OpeningGBV**: Sum of all cohorts' opening balances
- **InterestRevenue**: Sum of all cohorts' interest revenue
- **ClosingGBV**: Opening + Interest Revenue - Collections - Write-offs
- **Total_Coverage_Ratio**: Sum(Provision_Balance) / Sum(ClosingGBV)
- **Net_Impairment**: Impairment charge for the month
- **ClosingNBV**: ClosingGBV - Net_Impairment

---

## Part 2: Forecast_Details.xlsx Output

### Sheet: All_Cohorts

This sheet contains complete forecast for all Segment × Cohort combinations.

**Example Data (truncated for space):**

| ForecastMonth | Segment | Cohort | MOB | OpeningGBV | Coll_Principal_Rate | Coll_Principal_Approach | Coll_Interest_Rate | Coll_Interest_Approach | InterestRevenue_Rate | InterestRevenue_Approach | ... | ClosingGBV | Total_Coverage_Ratio | Total_Provision_Balance | Net_Impairment | ClosingNBV |
|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|
| 2025-10-31 | NRP-S | 202001 | 69 | 4,571.87 | -0.150000 | CohortAvg | -0.007374 | CohortTrend | 0.100000 | CohortTrend | ... | 3,890.48 | 0.1200 | 466.86 | 466.86 | 3,423.62 |
| 2025-10-31 | NON PRIME | 202001 | 69 | 86,444.65 | -0.050000 | Manual | 0.000000 | CohortTrend | 0.100000 | CohortTrend | ... | 82,842.79 | 0.1250 | 10,355.35 | 10,355.35 | 72,487.44 |
| 2025-11-30 | NRP-S | 202001 | 70 | 3,890.48 | -0.150000 | CohortAvg | -0.006774 | CohortTrend | 0.100000 | CohortTrend | ... | 3,312.96 | 0.1200 | 397.56 | 397.56 | 2,915.40 |

**Columns:**
- **ForecastMonth**: End of month date
- **Segment**: Loan segment
- **Cohort**: Origination cohort
- **MOB**: Months on book
- **OpeningGBV**: Opening balance
- **[Metric]_Rate**: Rate for each metric
- **[Metric]_Approach**: Calculation approach for each metric
- **[Metric]**: Calculated amount for each metric
- **ClosingGBV**: Closing balance
- **Total_Coverage_Ratio**: Provision coverage ratio
- **Total_Provision_Balance**: Provision balance
- **Net_Impairment**: Impairment charge
- **ClosingNBV**: Net book value

---

## Part 3: Impairment_Analysis.xlsx Output

### Sheet: Impairment_Detail

Detailed impairment calculations by month and cohort.

**Example Data:**

| ForecastMonth | Segment | Cohort | MOB | ClosingGBV | Total_Coverage_Ratio | Total_Provision_Balance | Total_Provision_Movement | Debt_Sale_WriteOffs | Debt_Sale_Coverage_Ratio | Debt_Sale_Provision_Release | Debt_Sale_Proceeds | Non_DS_Provision_Movement | Gross_Impairment_ExcludingDS | Debt_Sale_Impact | Net_Impairment |
|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|
| 2025-10-31 | NRP-S | 202001 | 69 | 3,890.48 | 0.1200 | 466.86 | 466.86 | 0.00 | 0.0000 | 0.00 | 0.00 | 466.86 | 466.86 | 0.00 | 466.86 |
| 2025-11-30 | NRP-S | 202001 | 70 | 3,312.96 | 0.1200 | 397.56 | -69.30 | 0.00 | 0.0000 | 0.00 | 0.00 | -69.30 | -69.30 | 0.00 | -69.30 |
| 2025-12-31 | NRP-S | 202001 | 71 | 2,823.02 | 0.1200 | 338.76 | -58.80 | 0.00 | 0.0000 | 0.00 | 0.00 | -58.80 | -58.80 | 0.00 | -58.80 |

**Key Metrics:**
- **Total_Provision_Movement**: Change in provision balance month-over-month
- **Debt_Sale_Provision_Release**: Provision released on debt sales
- **Non_DS_Provision_Movement**: Provision movement excluding debt sales
- **Gross_Impairment_ExcludingDS**: Provision movement + WO_Other
- **Debt_Sale_Impact**: DS_WriteOffs + DS_Provision_Release + DS_Proceeds
- **Net_Impairment**: Total impairment charge (Gross + DS Impact)

### Sheet: Coverage_Ratios

Coverage ratio trends by Segment and Cohort.

**Example Data:**

| Segment | Cohort | MOB | Total_Coverage_Ratio | Debt_Sale_Coverage_Ratio | Debt_Sale_Proceeds_Rate |
|---|---|---|---|---|---|
| NRP-S | 202001 | 69 | 0.1200 | N/A | N/A |
| NRP-S | 202001 | 70 | 0.1200 | N/A | N/A |
| NRP-S | 202001 | 71 | 0.1200 | N/A | N/A |
| NON PRIME | 202001 | 69 | 0.1250 | N/A | N/A |
| NON PRIME | 202504 | 1 | 0.1500 | 0.8500 | 0.9000 |

---

## Part 4: Validation_Report.xlsx Output

### Sheet: Reconciliation

Month-by-month reconciliation checks.

**Example Data:**

| ForecastMonth | Segment | Cohort | OpeningGBV | +InterestRevenue | -Coll_Principal | -Coll_Interest | -WO_DebtSold | -WO_Other | =ClosingGBV_Calculated | ClosingGBV_Forecast | Variance | Status |
|---|---|---|---|---|---|---|---|---|---|---|---|---|
| 2025-10-31 | NRP-S | 202001 | 4,571.87 | 38.10 | -685.78 | -33.71 | 0.00 | 0.00 | 3,890.48 | 3,890.48 | 0.00 | PASS |
| 2025-10-31 | NON PRIME | 202001 | 86,444.65 | 720.37 | -4,322.23 | 0.00 | 0.00 | 0.00 | 82,842.79 | 82,842.79 | 0.00 | PASS |
| 2025-11-30 | NRP-S | 202001 | 3,890.48 | 32.42 | -583.57 | -26.37 | 0.00 | 0.00 | 3,312.96 | 3,312.96 | 0.00 | PASS |

**Validation Rules:**
- ClosingGBV_Calculated = OpeningGBV + InterestRevenue - Coll_Principal - Coll_Interest - WO_DebtSold - WO_Other
- Status = PASS if |Variance| < 0.01, else FAIL

### Sheet: Validation_Checks

Summary of validation issues.

**Example Data:**

| Check | Total_Rows | Passed | Failed | Pass_Rate | Status |
|---|---|---|---|---|---|
| GBV_Reconciliation | 285 | 285 | 0 | 100.0% | PASS |
| NBV_Reconciliation | 285 | 285 | 0 | 100.0% | PASS |
| No_NaN_Values | 285 | 285 | 0 | 100.0% | PASS |
| No_Infinite_Values | 285 | 285 | 0 | 100.0% | PASS |
| Forecast_Chain_Continuity | 285 | 285 | 0 | 100.0% | PASS |
| Coverage_Ratio_Range | 285 | 285 | 0 | 100.0% | PASS |
| Overall | - | - | - | 100.0% | PASS |

---

## Part 5: Data Quality Checks

### 5.1 Input Data Validation

```
Check: Fact_Raw_Full.csv
  ✓ File exists
  ✓ 7,239 rows loaded
  ✓ 21 columns present
  ✓ CalendarMonth: 69 unique months (2020-01 to 2025-09)
  ✓ Cohort: 23 unique cohorts
  ✓ Segment: 5 unique segments (NON PRIME, NRP-S, NRP-M, NRP-L, PRIME)
  ✓ MOB: Range 0-68
  ✓ No missing required columns
  ✓ All numeric columns are numeric type
  ✓ No negative GBV (warning if found)
  
Check: Rate_Methodology.csv
  ✓ File exists
  ✓ 88 rules loaded
  ✓ All Approach values recognized
  ✓ MOB ranges valid (Start <= End)
  ✓ Param1 values valid for each approach
```

### 5.2 Forecast Validation

```
Check: Forecast Output
  ✓ 285 rows generated (95 Segment×Cohort × 3 months)
  ✓ All required columns present
  ✓ No NaN values in key columns
  ✓ No infinite values
  ✓ GBV reconciliation: 285/285 PASS (0.00% variance)
  ✓ NBV reconciliation: 285/285 PASS (0.00% variance)
  ✓ Forecast chain continuity: 285/285 PASS
  ✓ Coverage ratios in range [0.05, 0.50]: 285/285 PASS
  ✓ Rates within caps: 285/285 PASS
  
Overall Status: ALL CHECKS PASSED ✓
```

---

## Part 6: Key Validation Formulas

### 6.1 GBV Reconciliation

**Formula:**
```
ClosingGBV_Calculated = OpeningGBV 
                      + InterestRevenue 
                      - Coll_Principal 
                      - Coll_Interest 
                      - WO_DebtSold 
                      - WO_Other

Variance = |ClosingGBV_Calculated - ClosingGBV_Forecast|
Status = PASS if Variance < 0.01, else FAIL
```

**Example:**
```
OpeningGBV:        4,571.87
+ InterestRevenue:    38.10
- Coll_Principal:   -685.78
- Coll_Interest:     -33.71
- WO_DebtSold:        0.00
- WO_Other:           0.00
= ClosingGBV:      3,890.48 ✓
```

### 6.2 NBV Reconciliation

**Formula:**
```
ClosingNBV_Calculated = ClosingGBV - Net_Impairment

Variance = |ClosingNBV_Calculated - ClosingNBV_Forecast|
Status = PASS if Variance < 0.01, else FAIL
```

**Example:**
```
ClosingGBV:      3,890.48
- Net_Impairment:  466.86
= ClosingNBV:    3,423.62 ✓
```

### 6.3 Forecast Chain Continuity

**Formula:**
```
For each (Segment, Cohort):
  ClosingGBV[t] should equal OpeningGBV[t+1]
  
Variance = |ClosingGBV[t] - OpeningGBV[t+1]|
Status = PASS if Variance < 0.01 for all rows, else FAIL
```

**Example:**
```
Month 1: ClosingGBV = 3,890.48
Month 2: OpeningGBV = 3,890.48 ✓ (Continuous)
```

### 6.4 Impairment Calculation

**Formula:**
```
Total_Provision_Balance = ClosingGBV × Total_Coverage_Ratio
Total_Provision_Movement = Total_Provision_Balance[t] - Total_Provision_Balance[t-1]
Non_DS_Provision_Movement = Total_Provision_Movement + Debt_Sale_Provision_Release
Gross_Impairment_ExcludingDS = Non_DS_Provision_Movement + WO_Other
Debt_Sale_Impact = Debt_Sale_WriteOffs + Debt_Sale_Provision_Release + Debt_Sale_Proceeds
Net_Impairment = Gross_Impairment_ExcludingDS + Debt_Sale_Impact
```

**Example:**
```
ClosingGBV:                    3,890.48
× Total_Coverage_Ratio:           0.12
= Total_Provision_Balance:       466.86

Prior Provision Balance:         0.00
Total_Provision_Movement:      466.86
+ Debt_Sale_Provision_Release:   0.00
= Non_DS_Provision_Movement:   466.86
+ WO_Other:                      0.00
= Gross_Impairment_ExcludingDS: 466.86
+ Debt_Sale_Impact:              0.00
= Net_Impairment:              466.86 ✓
```

---

## Part 7: Common Issues & Troubleshooting

### 7.1 GBV Reconciliation Failures

**Issue:** ClosingGBV doesn't match calculated value

**Causes:**
- Rounding errors in rate calculations
- Missing or incorrect rate data
- Formula error in amount calculation

**Resolution:**
- Check rate precision (should be 6 decimals)
- Verify all rates are applied
- Review formula in spreadsheet

### 7.2 NaN or Infinite Values

**Issue:** Output contains NaN or Inf values

**Causes:**
- Division by zero (OpeningGBV = 0)
- Missing rate data
- Invalid rate values

**Resolution:**
- Check for zero opening balances
- Verify methodology rules cover all MOB ranges
- Validate rate values are within reasonable ranges

### 7.3 Forecast Chain Breaks

**Issue:** ClosingGBV[t] ≠ OpeningGBV[t+1]

**Causes:**
- Rounding in GBV calculation
- Missing months in forecast
- Data type conversion issues

**Resolution:**
- Use consistent decimal precision
- Ensure all months are included
- Check data types (float, not string)

### 7.4 Coverage Ratio Out of Range

**Issue:** Coverage ratio < 0.05 or > 0.50

**Causes:**
- Provision balance too high or low
- GBV calculation error
- Invalid coverage ratio assumption

**Resolution:**
- Review provision balance calculation
- Verify GBV reconciliation
- Check coverage ratio assumptions in methodology

---

## Part 8: Performance Benchmarks

Expected performance on full dataset (7,239 rows, 95 cohorts, 12 months):

| Step | Time (seconds) | Notes |
|---|---|---|
| Load data | 0.5 | Reading CSV |
| Calculate curves | 1.0 | Aggregation & rate calc |
| Extend curves | 0.5 | Creating extended rates |
| Calculate impairment | 1.5 | Provision calculations |
| Generate seed | 0.2 | Filtering last month |
| Build rate lookup | 5.0 | Methodology lookups |
| Build impairment lookup | 3.0 | Coverage ratio lookups |
| Run forecast | 2.0 | 12 months × 95 cohorts |
| Generate outputs | 1.0 | DataFrame aggregation |
| Export to Excel | 2.0 | Writing 4 workbooks |
| **Total** | **~17 seconds** | End-to-end |

---

## Part 9: Sample Test Cases

### Test Case 1: Single Cohort, Single Month

**Input:**
- Fact_Raw: 1 cohort (202001), 1 segment (NRP-S), 1 month
- Forecast: 1 month

**Expected Output:**
- 1 forecast row
- ClosingGBV = OpeningGBV + InterestRevenue - Collections - WriteOffs
- ClosingNBV = ClosingGBV - Net_Impairment

### Test Case 2: Multiple Cohorts, Multiple Months

**Input:**
- Fact_Raw: 5 cohorts, 5 segments, 69 months
- Forecast: 12 months

**Expected Output:**
- 95 × 12 = 1,140 forecast rows (if all cohorts active)
- All validation checks pass
- No NaN or infinite values

### Test Case 3: Debt Sale Month

**Input:**
- Debt_Sale_Schedule: 1 debt sale in month 2
- Debt_Sale_WriteOffs: 1,000
- Debt_Sale_Coverage_Ratio: 0.85
- Debt_Sale_Proceeds_Rate: 0.90

**Expected Output:**
- Month 2: Debt_Sale_Provision_Release = 850, Debt_Sale_Proceeds = 900
- Debt_Sale_Impact = 1,000 + 850 + 900 = 2,750
- Net_Impairment includes debt sale impact

---

## Part 10: Acceptance Criteria

The model is considered complete when:

- [ ] All 4 Excel workbooks are generated
- [ ] Forecast_Summary shows correct totals by month and segment
- [ ] Forecast_Details shows all cohorts with correct calculations
- [ ] Impairment_Analysis shows correct provision movements
- [ ] Validation_Report shows 100% pass rate on all checks
- [ ] GBV reconciliation variance < 0.01 for all rows
- [ ] NBV reconciliation variance < 0.01 for all rows
- [ ] No NaN or infinite values in output
- [ ] Forecast chain continuity verified
- [ ] Coverage ratios within [0.05, 0.50] range
- [ ] All rates within caps
- [ ] Performance < 30 seconds for full dataset
