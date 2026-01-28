# Complete Backbook Forecasting Model - Documentation Index

## Quick Navigation

### 📋 Start Here
- **README.md** - Overview and how to use this documentation package

### 🏗️ Architecture & Design
- **01_SYSTEM_DESIGN.md** - Complete system architecture, data flow, and calculation specifications

### 💻 Implementation
- **02_IMPLEMENTATION_GUIDE.md** - Step-by-step implementation guide for Claude Code with detailed function specifications

### ✅ Validation & Testing
- **03_EXAMPLE_OUTPUTS.md** - Expected output formats, validation checks, and troubleshooting

### ⚡ Quick Reference
- **04_QUICK_REFERENCE.md** - Cheat sheet with key formulas, approaches, and standards

---

## Document Overview

### README.md (11 KB, 365 lines)
**Purpose:** Overview and navigation guide

**Key Sections:**
- Documentation structure
- How to use this documentation
- Key concepts
- Input/output files
- Key formulas
- Validation checks
- Performance targets
- Common pitfalls
- Deliverables checklist

**Read When:** Starting the project or need overview

---

### 01_SYSTEM_DESIGN.md (16 KB, 452 lines)
**Purpose:** Complete system architecture and design specifications

**Key Sections:**
1. Model Architecture (data flow, components)
2. Detailed Calculations
   - Collections & Interest Revenue
   - Closing GBV
   - Impairment & Provision (Actuals & Forecast)
   - Closing NBV
3. Input Data Requirements
4. Output Structure (4 Excel workbooks)
5. Rate Calculation Approaches (6 types)
6. Rate Caps (per metric)
7. Methodology Lookup Specificity Scoring
8. Forecast Loop Logic
9. Data Type & Format Standards
10. Error Handling & Validation
11. Configuration Parameters
12. Next Steps

**Read When:** Need to understand system design or calculation logic

---

### 02_IMPLEMENTATION_GUIDE.md (27 KB, 868 lines)
**Purpose:** Detailed implementation specifications for Claude Code

**Key Sections:**
1. Project Structure
2. Module Structure & Functions (13 functional areas)
3. Detailed Function Specifications
   - Data Loading (3 functions)
   - Curves Calculation (2 functions)
   - Impairment Curves (2 NEW functions)
   - Seed Generation (2 functions)
   - Methodology Lookup (2 functions)
   - Rate Calculation (4 functions)
   - Rate Application (2 functions)
   - Forecast Engine (2 functions)
   - Output Generation (5 functions)
   - Main Orchestration (1 function)
4. Command Line Interface
5. Error Handling & Logging
6. Testing & Validation
7. Dependencies & Requirements
8. Code Quality Standards
9. Deliverables Checklist
10. Known Limitations & Future Enhancements

**Read When:** Implementing the Python module - follow function by function

---

### 03_EXAMPLE_OUTPUTS.md (14 KB, 410 lines)
**Purpose:** Expected output formats and validation specifications

**Key Sections:**
1. Forecast_Summary.xlsx Output (example data)
2. Forecast_Details.xlsx Output (example data)
3. Impairment_Analysis.xlsx Output (2 sheets)
4. Validation_Report.xlsx Output (2 sheets)
5. Data Quality Checks
6. Forecast Validation
7. Key Validation Formulas
8. Common Issues & Troubleshooting
9. Performance Benchmarks
10. Sample Test Cases
11. Acceptance Criteria

**Read When:** Validating outputs, testing, or troubleshooting

---

### 04_QUICK_REFERENCE.md (8.6 KB, 321 lines)
**Purpose:** Quick lookup reference for key concepts and standards

**Key Sections:**
- Key Formulas at a Glance
- Rate Calculation Approaches (table)
- Rate Caps (table)
- Input Files (table)
- Output Files (table)
- Function Call Sequence
- Methodology Lookup Specificity Scoring
- Validation Checks
- Command Line Usage
- Common Pitfalls to Avoid
- Data Type Standards (table)
- Performance Targets (table)
- Key Metrics to Track
- Error Handling (code examples)
- Testing Checklist
- Deliverables

**Read When:** Need a quick reference for a specific formula or standard

---

## How to Use This Documentation

### For Claude Code Implementation

**Step 1: Understand the Concepts (30 minutes)**
- Read README.md for overview
- Skim 04_QUICK_REFERENCE.md for key concepts

**Step 2: Learn the Architecture (1 hour)**
- Read 01_SYSTEM_DESIGN.md sections 1-7
- Understand data flow and calculation logic

**Step 3: Implement the Module (4-6 hours)**
- Follow 02_IMPLEMENTATION_GUIDE.md
- Implement each function according to specifications
- Use 04_QUICK_REFERENCE.md for quick lookups

**Step 4: Test and Validate (2-3 hours)**
- Use 03_EXAMPLE_OUTPUTS.md for validation checks
- Run unit and integration tests
- Verify outputs match expected formats

**Total Estimated Time: 7-10 hours**

---

### For Validation & Testing

**Step 1: Understand Expected Outputs**
- Review 03_EXAMPLE_OUTPUTS.md sections 1-4

**Step 2: Run Validation Checks**
- Follow validation formulas in 03_EXAMPLE_OUTPUTS.md section 6
- Use reconciliation checks in 03_EXAMPLE_OUTPUTS.md section 4

**Step 3: Troubleshoot Issues**
- Check 03_EXAMPLE_OUTPUTS.md section 8 (Common Issues)
- Refer to 04_QUICK_REFERENCE.md (Common Pitfalls)

**Step 4: Verify Acceptance Criteria**
- Follow checklist in 03_EXAMPLE_OUTPUTS.md section 10

---

### For Troubleshooting

**Issue: Model doesn't run**
1. Check 04_QUICK_REFERENCE.md - Error Handling section
2. Review 02_IMPLEMENTATION_GUIDE.md - Error Handling & Logging section
3. Verify input files exist and have correct format

**Issue: GBV doesn't reconcile**
1. Check 03_EXAMPLE_OUTPUTS.md - GBV Reconciliation formula
2. Verify all rates are applied correctly
3. Check for rounding issues

**Issue: Outputs don't match expected format**
1. Review 03_EXAMPLE_OUTPUTS.md sections 1-4
2. Check data types in 04_QUICK_REFERENCE.md - Data Type Standards
3. Verify Excel formatting in 02_IMPLEMENTATION_GUIDE.md - export_to_excel function

---

## Key Metrics Summary

### Input Data
- **Fact_Raw_Full.csv:** 7,239 rows, 21 columns, 69 months, 23 cohorts
- **Rate_Methodology.csv:** 88 rules
- **Debt_Sale_Schedule.csv:** Variable (debt sale months only)

### Output Data
- **Forecast_Summary.xlsx:** By month & segment
- **Forecast_Details.xlsx:** All cohorts (95 × 12 = 1,140 rows for 12 months)
- **Impairment_Analysis.xlsx:** Impairment detail
- **Validation_Report.xlsx:** Validation checks

### Performance
- Total runtime: < 30 seconds
- Load data: < 1 second
- Calculate curves: < 2 seconds
- Build lookups: < 10 seconds
- Run forecast: < 5 seconds
- Generate outputs: < 5 seconds

---

## Key Formulas Summary

### Collections & Interest
```
Coll_Principal = OpeningGBV × Coll_Principal_Rate
Coll_Interest = OpeningGBV × Coll_Interest_Rate
InterestRevenue = OpeningGBV × InterestRevenue_Rate / 12
```

### Closing GBV
```
ClosingGBV = OpeningGBV + InterestRevenue - Coll_Principal - Coll_Interest - WO_DebtSold - WO_Other
```

### Impairment
```
Total_Provision_Balance = ClosingGBV × Total_Coverage_Ratio
Net_Impairment = (Total_Provision_Movement + Debt_Sale_Provision_Release + WO_Other) + (Debt_Sale_WriteOffs + Debt_Sale_Provision_Release + Debt_Sale_Proceeds)
```

### Closing NBV
```
ClosingNBV = ClosingGBV - Net_Impairment
```

---

## Rate Calculation Approaches

| Approach | Formula | Use Case |
|----------|---------|----------|
| CohortAvg | Average(last N MOBs) | Stable trends |
| CohortTrend | Linear regression | Trending data |
| DonorCohort | Copy from donor | New cohorts |
| SegMedian | Median across segment | Cross-cohort |
| Manual | Fixed override | Assumptions |
| Zero | Force to zero | No activity |

---

## Validation Checks

All outputs must pass:
1. ✓ GBV Reconciliation (variance < 0.01)
2. ✓ NBV Reconciliation (variance < 0.01)
3. ✓ No NaN or Infinite Values
4. ✓ Forecast Chain Continuity
5. ✓ Coverage Ratios in Range [0.05, 0.50]
6. ✓ Rates Within Caps

---

## File Statistics

| File | Size | Lines | Purpose |
|------|------|-------|---------|
| README.md | 11 KB | 365 | Overview & navigation |
| 01_SYSTEM_DESIGN.md | 16 KB | 452 | Architecture & design |
| 02_IMPLEMENTATION_GUIDE.md | 27 KB | 868 | Implementation specs |
| 03_EXAMPLE_OUTPUTS.md | 14 KB | 410 | Validation & examples |
| 04_QUICK_REFERENCE.md | 8.6 KB | 321 | Quick reference |
| **Total** | **76.6 KB** | **2,416** | Complete documentation |

---

## Recommended Reading Order

1. **README.md** (5 min) - Get oriented
2. **04_QUICK_REFERENCE.md** (10 min) - Learn key concepts
3. **01_SYSTEM_DESIGN.md** (30 min) - Understand architecture
4. **02_IMPLEMENTATION_GUIDE.md** (60 min) - Learn implementation details
5. **03_EXAMPLE_OUTPUTS.md** (20 min) - Understand validation

**Total: ~2 hours** for complete understanding

---

## Deliverables Checklist

- [ ] backbook_forecast.py (single module)
- [ ] requirements.txt (dependencies)
- [ ] README.md (usage instructions)
- [ ] Forecast_Summary.xlsx (output)
- [ ] Forecast_Details.xlsx (output)
- [ ] Impairment_Analysis.xlsx (output)
- [ ] Validation_Report.xlsx (output)
- [ ] Unit tests
- [ ] Integration tests
- [ ] Error handling & logging
- [ ] Command line interface

---

## Next Steps

1. **Review README.md** - Get overview
2. **Read 04_QUICK_REFERENCE.md** - Learn key concepts
3. **Study 01_SYSTEM_DESIGN.md** - Understand architecture
4. **Follow 02_IMPLEMENTATION_GUIDE.md** - Build the module
5. **Use 03_EXAMPLE_OUTPUTS.md** - Validate outputs

---

## Support

For questions about:
- **Architecture:** See 01_SYSTEM_DESIGN.md
- **Implementation:** See 02_IMPLEMENTATION_GUIDE.md
- **Validation:** See 03_EXAMPLE_OUTPUTS.md
- **Quick lookup:** See 04_QUICK_REFERENCE.md
- **Overview:** See README.md

---

**Documentation Version:** 1.0  
**Created:** 2026-01-13  
**Status:** Complete & Ready for Claude Code Implementation
