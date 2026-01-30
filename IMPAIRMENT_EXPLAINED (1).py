#!/usr/bin/env python3
"""
IMPAIRMENT CALCULATION EXPLAINED

This script provides a detailed walkthrough of how impairment is calculated
in the backbook forecast model, with a step-by-step example.

Updated: Now uses WO_DebtSold (forecast from rates) as Debt Sale WriteOffs
with a fixed 78.5% DS Coverage Ratio.
"""

print("""
================================================================================
IMPAIRMENT CALCULATION - COMPLETE EXPLANATION
================================================================================

The model calculates several impairment metrics. Here's what each means:

--------------------------------------------------------------------------------
KEY CONCEPTS
--------------------------------------------------------------------------------

1. PROVISION BALANCE (aka Allowance for Credit Losses)
   - The reserve set aside to cover expected future losses
   - Calculated as: ClosingGBV × Coverage_Ratio
   - Example: £100,000 GBV × 12% coverage = £12,000 provision

2. COVERAGE RATIO
   - What % of GBV is covered by provisions
   - Higher = more conservative/risky portfolio
   - Typical range: 5% (prime) to 50%+ (subprime)

3. PROVISION MOVEMENT
   - Change in provision balance from prior month
   - Positive = increased provisions (expense)
   - Negative = released provisions (income)

4. DEBT SALE (when you sell bad loans to a third party)
   - WO_DebtSold (from rates) IS the Debt_Sale_WriteOffs
   - DS_Coverage_Ratio = 78.5% FIXED for all cohorts/segments
   - DS_Provision_For_Pool: Provision allocated to debt sale pool
   - Debt_Sale_Provision_Release: Provision freed up (was covering those loans)
   - Debt_Sale_Proceeds: Cash received from buyer

5. CORE VALUES (after removing debt sale portion)
   - Core_Provision: Remaining provision after DS provision removed
   - Core_GBV: Remaining GBV after DS writeoffs removed
   - Core_Coverage_Ratio: Coverage on the remaining "good" loans

--------------------------------------------------------------------------------
THE IMPAIRMENT FORMULAS (UPDATED)
--------------------------------------------------------------------------------

STEP 1: Forecast WO_DebtSold from Rates
    WO_DebtSold = OpeningGBV × WO_DebtSold_Rate
    This IS the Debt Sale WriteOffs

STEP 2: Calculate Debt Sale Provision for Pool
    DS_Coverage_Ratio = 78.5% (FIXED)
    DS_Provision_For_Pool = DS_Coverage_Ratio × WO_DebtSold
    This is the provision that was covering the loans being sold

STEP 3: Calculate Core Values (after debt sale)
    Core_Provision = Prior_Provision - DS_Provision_For_Pool
    Core_GBV = OpeningGBV - WO_DebtSold
    Core_Coverage_Ratio = Core_Provision / Core_GBV

STEP 4: Calculate ClosingGBV
    ClosingGBV = OpeningGBV + InterestRevenue - Collections - WriteOffs
    (WriteOffs include WO_DebtSold and WO_Other)

STEP 5: Calculate New Provision Balance
    Provision_Balance = ClosingGBV × Coverage_Ratio (from methodology)

STEP 6: Calculate Provision Movement
    Provision_Movement = Provision_Balance[t] - Provision_Balance[t-1]

STEP 7: Calculate Debt Sale components
    Debt_Sale_Provision_Release = DS_Provision_For_Pool
    Debt_Sale_Proceeds = Debt_Sale_Proceeds_Rate × WO_DebtSold

STEP 8: Calculate Net Impairment
    # Sign convention: DS_Provision_Release and DS_Proceeds stored as NEGATIVE (credit)
    Non_DS_Provision_Movement = Provision_Movement - Debt_Sale_Provision_Release
    Gross_Impairment_ExcludingDS = Non_DS_Provision_Movement + WO_Other
    Debt_Sale_Impact = WO_DebtSold + Debt_Sale_Provision_Release + Debt_Sale_Proceeds
    Net_Impairment = Gross_Impairment_ExcludingDS + Debt_Sale_Impact

STEP 9: Closing NBV
    ClosingNBV = ClosingGBV - Net_Impairment

================================================================================
END-TO-END EXAMPLE: Forecasting with Debt Sale
================================================================================

Starting Point (from last actual month):
    OpeningGBV = £100,000
    Prior_Provision_Balance = £12,000 (12% coverage)
    WO_DebtSold_Rate = 2% (from rate methodology)
    Coverage_Ratio = 12.5% (from methodology)
    DS_Coverage_Ratio = 78.5% (FIXED)
    DS_Proceeds_Rate = 90%

STEP 1: Calculate WO_DebtSold (which IS Debt Sale WriteOffs)
    WO_DebtSold = £100,000 × 2% = £2,000
    Debt_Sale_WriteOffs = £2,000

STEP 2: Calculate DS Provision for Pool
    DS_Provision_For_Pool = 78.5% × £2,000 = £1,570

STEP 3: Calculate Core Values
    Core_Provision = £12,000 - £1,570 = £10,430
    Core_GBV = £100,000 - £2,000 = £98,000
    Core_Coverage_Ratio = £10,430 / £98,000 = 10.64%

STEP 4: Calculate ClosingGBV
    Assume: InterestRevenue = £800, Collections = £3,000, WO_Other = £200
    ClosingGBV = £100,000 + £800 - £3,000 - £2,000 - £200 = £95,600

STEP 5: Calculate New Provision Balance
    Provision_Balance = £95,600 × 12.5% = £11,950

STEP 6: Calculate Provision Movement
    Provision_Movement = £11,950 - £12,000 = -£50 (small release)

STEP 7: Calculate Debt Sale components (stored as NEGATIVE - credit convention)
    Debt_Sale_Provision_Release = -£1,570 (credit: release from provision)
    Debt_Sale_Proceeds = -(90% × £2,000) = -£1,800 (credit: cash inflow)

STEP 8: Calculate Net Impairment
    # Formula uses minus: Non_DS = Total - DS_Release (where DS_Release is negative)
    Non_DS_Provision_Movement = -£50 - (-£1,570) = -£50 + £1,570 = £1,520
    Gross_Impairment_ExcludingDS = £1,520 + £200 = £1,720
    # Debt Sale Impact with credit convention negatives:
    Debt_Sale_Impact = £2,000 + (-£1,570) + (-£1,800) = -£1,370 (net gain from sale)
    Net_Impairment = £1,720 + (-£1,370) = £350

STEP 9: Calculate Closing NBV
    ClosingNBV = £95,600 - £350 = £95,250

================================================================================
SUMMARY OF KEY CHANGES
================================================================================

1. WO_DebtSold (forecast from rates) IS the Debt Sale WriteOffs
   - No separate Debt_Sale_Schedule file needed for debt sale amounts
   - The WO_DebtSold_Rate in Rate_Methodology controls debt sale volume

2. DS Coverage Ratio = 78.5% FIXED
   - Applied to all cohorts, segments, and forecast months
   - Configurable in Config.DS_COVERAGE_RATIO

3. New Metrics Added:
   - DS_Provision_For_Pool: DS_Coverage × WO_DebtSold
   - Core_Provision: Prior provision minus DS provision for pool
   - Core_GBV: Opening GBV minus debt sale writeoffs
   - Core_Coverage_Ratio: Core provision / Core GBV

4. The Core values show what remains AFTER the debt sale:
   - Core_Provision = provision covering remaining loans
   - Core_GBV = GBV of remaining loans
   - Core_Coverage_Ratio = how well covered the remaining loans are

================================================================================
CONFIGURING WO_DebtSold RATES
================================================================================

To control how much is forecast as debt sale writeoffs, update the
Rate_Methodology.csv file with appropriate WO_DebtSold rate rules:

Example: To forecast 2% per month as debt sales for all cohorts:
    Segment,Cohort,MOB_Start,MOB_End,Metric,Approach,Param1,Param2
    ALL,ALL,1,999,WO_DebtSold,Manual,0.02,

Example: To use cohort average for debt sale rates:
    Segment,Cohort,MOB_Start,MOB_End,Metric,Approach,Param1,Param2
    ALL,ALL,1,999,WO_DebtSold,CohortAvg,6,

Example: No debt sales (Zero approach):
    Segment,Cohort,MOB_Start,MOB_End,Metric,Approach,Param1,Param2
    ALL,ALL,1,999,WO_DebtSold,Zero,,

================================================================================
""")
