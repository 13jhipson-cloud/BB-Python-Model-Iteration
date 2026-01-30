# Complete Backbook Forecasting Model - Documentation Package

## Overview

This documentation package contains comprehensive specifications for building a complete Python backbook forecasting model that calculates loan portfolio performance including collections, GBV, impairment, and NBV over 12-36 months.

The model is designed to be built by Claude Code and should produce identical outputs to the existing Power Query implementation while running significantly faster.

---

## Documentation Structure

### 1. **01_SYSTEM_DESIGN.md** (Primary Reference)
**Purpose:** Architecture and design specifications

**Contents:**
- Model architecture and data flow
- Detailed calculation formulas
- Input data requirements
- Output structure
- Rate calculation approaches
- Rate caps and methodology lookup
- Forecast loop logic
- Configuration parameters

**Use When:** You need to understand the overall system design, data flow, or calculation logic.

---

### 2. **02_IMPLEMENTATION_GUIDE.md** (Implementation Blueprint)
**Purpose:** Step-by-step implementation specifications for Claude Code

**Contents:**
- Project structure
- Module structure and function signatures
- Detailed function specifications with pseudocode
- Data loading functions
- Curves calculation functions
- Impairment calculations (NEW)
- Seed generation functions
- Methodology lookup functions
- Rate calculation functions
- Forecast engine functions
- Output generation functions
- Command line interface
- Error handling and logging
- Testing specifications
- Dependencies and requirements

**Use When:** Building the Python module. Follow this guide function-by-function.

---

### 3. **03_EXAMPLE_OUTPUTS.md** (Validation & Testing)
**Purpose:** Expected output formats and validation specifications

**Contents:**
- Example output data for each Excel workbook
- Data quality checks
- Validation formulas and checks
- Common issues and troubleshooting
- Performance benchmarks
- Sample test cases
- Acceptance criteria

**Use When:** Validating outputs, testing the model, or troubleshooting issues.

---

### 4. **04_QUICK_REFERENCE.md** (Cheat Sheet)
**Purpose:** Quick lookup for key formulas, approaches, and standards

**Contents:**
- Key formulas at a glance
- Rate calculation approaches
- Rate caps
- Input/output file specifications
- Function call sequence
- Methodology lookup scoring
- Validation checks
- Command line usage
- Common pitfalls
- Data type standards
- Performance targets
- Error handling examples
- Testing checklist

**Use When:** You need a quick reference for a specific formula, approach, or standard.

---

## How to Use This Documentation

### For Claude Code Implementation

1. **Start with 04_QUICK_REFERENCE.md** - Get familiar with key concepts
2. **Read 01_SYSTEM_DESIGN.md** - Understand the overall architecture
3. **Follow 02_IMPLEMENTATION_GUIDE.md** - Implement function by function
4. **Validate with 03_EXAMPLE_OUTPUTS.md** - Test and verify outputs

### For Validation & Testing

1. **Review 03_EXAMPLE_OUTPUTS.md** - Understand expected outputs
2. **Run validation checks** - Verify GBV and NBV reconciliation
3. **Check 04_QUICK_REFERENCE.md** - Troubleshoot issues
4. **Compare with examples** - Ensure outputs match expected format

### For Troubleshooting

1. **Check 04_QUICK_REFERENCE.md** - Common pitfalls section
2. **Review 03_EXAMPLE_OUTPUTS.md** - Validation checks and troubleshooting
3. **Consult 01_SYSTEM_DESIGN.md** - Verify calculation logic
4. **Check 02_IMPLEMENTATION_GUIDE.md** - Review function specifications

---

## Key Concepts

### Collections & Interest Revenue
Calculated using historical rate curves applied to opening GBV. Rates are selected via methodology lookup with specificity scoring. Rates are capped to prevent unrealistic forecasts.

### Closing GBV
Calculated as: Opening GBV + Interest Revenue - Collections - Write-offs

### Impairment & Provisions
Calculated using coverage ratios applied to closing GBV. Includes special handling for debt sales. Tracks provision movements and impairment charges.

### Closing NBV
Calculated as: Closing GBV - Net Impairment

### Rate Calculation Approaches
- **CohortAvg:** Average of last N MOBs
- **CohortTrend:** Linear regression extrapolation
- **DonorCohort:** Copy from another cohort
- **SegMedian:** Median across segment
- **Manual:** Fixed override
- **Zero:** Force to zero

### Methodology Lookup
Selects the most specific matching rule using specificity scoring (Segment +8, Cohort +4, Metric +2, MOB range width as tiebreaker).

---

## Input Files

### Fact_Raw_Full.csv
Historical loan data with 7,239 rows and 21 columns including:
- CalendarMonth, Cohort, Segment, MOB
- OpeningGBV, collections, interest revenue, write-offs
- **NEW:** Provision_Balance, Debt_Sale_WriteOffs, Debt_Sale_Provision_Release, Debt_Sale_Proceeds

### Rate_Methodology.csv
Control table with 88 rules defining rate calculation approaches for each metric.

### Debt_Sale_Schedule.csv (Optional)
Debt sale assumptions by month, segment, and cohort.

---

## Output Files

| File | Purpose |
|------|---------|
| **Forecast_Summary.xlsx** | High-level summary by month and segment |
| **Forecast_Details.xlsx** | Complete forecast for all cohorts |
| **Impairment_Analysis.xlsx** | Impairment detail and coverage trends |
| **Validation_Report.xlsx** | Validation checks and reconciliation |

---

## Key Formulas

### GBV Reconciliation
```
ClosingGBV = OpeningGBV + InterestRevenue - Coll_Principal - Coll_Interest - WO_DebtSold - WO_Other
```

### NBV Reconciliation
```
ClosingNBV = ClosingGBV - Net_Impairment
```

### Impairment
```
Total_Provision_Balance = ClosingGBV × Total_Coverage_Ratio
Total_Provision_Movement = Total_Provision_Balance[t] - Total_Provision_Balance[t-1]

# Sign convention: DS_Provision_Release and DS_Proceeds stored as NEGATIVE (credit)
Non_DS_Provision_Movement = Total_Provision_Movement - Debt_Sale_Provision_Release
Gross_Impairment_ExcludingDS = Non_DS_Provision_Movement + WO_Other
Debt_Sale_Impact = Debt_Sale_WriteOffs + Debt_Sale_Provision_Release + Debt_Sale_Proceeds
Net_Impairment = Gross_Impairment_ExcludingDS + Debt_Sale_Impact
```

---

## Validation Checks

All outputs must pass these checks:

1. **GBV Reconciliation:** ClosingGBV = OpeningGBV + InterestRevenue - Collections - WriteOffs
2. **NBV Reconciliation:** ClosingNBV = ClosingGBV - Net_Impairment
3. **No NaN or Infinite Values**
4. **Forecast Chain Continuity:** ClosingGBV[t] = OpeningGBV[t+1]
5. **Coverage Ratios in Range:** [0.05, 0.50]
6. **Rates Within Caps**

---

## Performance Targets

- Load data: < 1 second
- Calculate curves: < 2 seconds
- Build lookups: < 10 seconds
- Run forecast: < 5 seconds
- Generate outputs: < 5 seconds
- **Total: < 30 seconds** for full dataset

---

## Common Pitfalls to Avoid

1. **Date Parsing:** Handle both M/D/YYYY and MM/DD/YYYY formats
2. **Cohort Type:** Keep cohorts as strings (not integers)
3. **Division by Zero:** Check OpeningGBV > 0 before dividing
4. **Rounding:** Use consistent decimal precision
5. **Forecast Chain:** Ensure ClosingGBV[t] = OpeningGBV[t+1]
6. **Rate Caps:** Apply AFTER rate calculation, BEFORE amount calculation
7. **Impairment:** Track provision balance separately from GBV
8. **Debt Sales:** Only apply debt sale logic in sale months

---

## Deliverables Checklist

- [ ] **backbook_forecast.py** - Single Python module with all code
- [ ] **requirements.txt** - Dependencies (pandas, numpy, openpyxl, python-dateutil)
- [ ] **README.md** - Usage instructions
- [ ] **4 Excel workbooks** - Forecast_Summary, Forecast_Details, Impairment_Analysis, Validation_Report
- [ ] **Unit tests** - Test individual functions
- [ ] **Integration tests** - Test end-to-end workflow
- [ ] **Error handling** - Graceful error messages
- [ ] **Logging** - Progress and debug information
- [ ] **Command line interface** - argparse for CLI arguments

---

## Example Usage

```bash
# Run forecast with all options
python backbook_forecast.py \
    --fact-raw Fact_Raw_Full.csv \
    --methodology Rate_Methodology.csv \
    --debt-sale Debt_Sale_Schedule.csv \
    --output output/ \
    --months 12

# Run forecast with defaults
python backbook_forecast.py \
    --fact-raw Fact_Raw_Full.csv \
    --methodology Rate_Methodology.csv

# Run tests
python -m pytest test_backbook_forecast.py -v
```

---

## Data Type Standards

| Field | Type | Format | Example |
|-------|------|--------|---------|
| ForecastMonth | datetime | YYYY-MM-DD | 2025-10-31 |
| Segment | string | Title case | NRP-S |
| Cohort | string | YYYYMM | 202001 |
| MOB | integer | Integer | 69 |
| OpeningGBV | float | 2 decimals | 4571.87 |
| Rate | float | 6 decimals | -0.150000 |
| Amount | float | 2 decimals | -685.78 |
| Coverage Ratio | float | 4 decimals | 0.1250 |

---

## Rate Caps

| Metric | Min | Max |
|--------|-----|-----|
| Coll_Principal | -0.15 | 0 |
| Coll_Interest | -0.10 | 0 |
| InterestRevenue | 0.10 | 0.50 |
| WO_DebtSold | 0 | 0.12 |
| WO_Other | 0 | 0.01 |
| ContraSettlements_Principal | -0.06 | 0 |
| ContraSettlements_Interest | -0.005 | 0 |
| Total_Coverage_Ratio | 0.05 | 0.50 |
| Debt_Sale_Coverage_Ratio | 0.50 | 1.00 |
| Debt_Sale_Proceeds_Rate | 0.30 | 1.00 |

---

## Support & Troubleshooting

### Issue: File Not Found
- Ensure all input files are in the same directory as the script
- Or provide full paths to files
- Check file names match exactly (case-sensitive on Linux/Mac)

### Issue: GBV Reconciliation Fails
- Check rate precision (should be 6 decimals)
- Verify all rates are applied
- Review formula: ClosingGBV = OpeningGBV + InterestRevenue - Collections - WriteOffs

### Issue: NaN or Infinite Values
- Check for zero opening balances
- Verify methodology rules cover all MOB ranges
- Validate rate values are within reasonable ranges

### Issue: Forecast Chain Breaks
- Check for rounding issues
- Ensure all months are included
- Verify data types (float, not string)

### Issue: Coverage Ratio Out of Range
- Review provision balance calculation
- Verify GBV reconciliation
- Check coverage ratio assumptions

---

## Next Steps

1. **Review Documentation:** Start with 04_QUICK_REFERENCE.md
2. **Understand Architecture:** Read 01_SYSTEM_DESIGN.md
3. **Implement Module:** Follow 02_IMPLEMENTATION_GUIDE.md
4. **Validate Outputs:** Use 03_EXAMPLE_OUTPUTS.md
5. **Test Thoroughly:** Run unit and integration tests
6. **Verify Results:** Check validation reports

---

## Version History

| Version | Date | Changes |
|---------|------|---------|
| 1.0 | 2026-01-13 | Initial documentation for complete backbook model with impairment |

---

## Document References

- **System Design:** 01_SYSTEM_DESIGN.md
- **Implementation Guide:** 02_IMPLEMENTATION_GUIDE.md
- **Example Outputs:** 03_EXAMPLE_OUTPUTS.md
- **Quick Reference:** 04_QUICK_REFERENCE.md

---

## License & Usage

This documentation is provided for internal use only. The model is designed to forecast loan portfolio performance and should be validated against actual results before use in production.

---

## Contact

For questions or clarifications regarding this documentation, please refer to the relevant section in the documentation files above.
