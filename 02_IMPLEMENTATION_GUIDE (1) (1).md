# Complete Backbook Forecasting Model - Implementation Guide for Claude Code

## Overview

This guide provides step-by-step instructions for implementing a complete Python backbook forecasting model that calculates collections, interest revenue, GBV, impairment, and NBV. The model should be built as a single, well-structured Python module that can be run from the command line.

---

## Part 1: Project Structure

```
backbook_forecast/
├── backbook_forecast.py          # Main module (all code)
├── requirements.txt              # Dependencies
├── Fact_Raw_Full.csv            # Input: historical data
├── Rate_Methodology.csv         # Input: rate calculation rules
├── Debt_Sale_Schedule.csv       # Input: debt sale assumptions (optional)
├── output/
│   ├── Forecast_Summary.xlsx    # High-level summary
│   ├── Forecast_Details.xlsx    # Complete forecast
│   ├── Impairment_Analysis.xlsx # Impairment detail
│   └── Validation_Report.xlsx   # Validation checks
└── README.md                     # Usage instructions
```

---

## Part 2: Module Structure & Functions

### 2.1 Core Modules

```python
# backbook_forecast.py

# 1. CONFIGURATION
class Config:
    - MAX_MONTHS
    - LOOKBACK_PERIODS
    - MOB_THRESHOLD
    - RATE_CAPS (dict)
    - IMPAIRMENT_CAPS (dict)
    - SEGMENTS (list)
    - METRICS (list)

# 2. HELPER FUNCTIONS
def parse_date(date_val) -> pd.Timestamp
def end_of_month(date: pd.Timestamp) -> pd.Timestamp
def clean_cohort(cohort_val) -> str
def safe_divide(numerator, denominator) -> float

# 3. DATA LOADING
def load_fact_raw(filepath: str) -> pd.DataFrame
def load_rate_methodology(filepath: str) -> pd.DataFrame
def load_debt_sale_schedule(filepath: str) -> pd.DataFrame

# 4. CURVES CALCULATION
def calculate_curves_base(fact_raw: pd.DataFrame) -> pd.DataFrame
def extend_curves(curves_base: pd.DataFrame, max_months: int) -> pd.DataFrame

# 5. IMPAIRMENT CURVES (NEW)
def calculate_impairment_actuals(fact_raw: pd.DataFrame) -> pd.DataFrame
def calculate_impairment_curves(impairment_actuals: pd.DataFrame) -> pd.DataFrame

# 6. SEED GENERATION
def generate_seed_curves(fact_raw: pd.DataFrame) -> pd.DataFrame
def generate_impairment_seed(fact_raw: pd.DataFrame) -> pd.DataFrame

# 7. METHODOLOGY LOOKUP
def get_methodology(methodology_df, segment, cohort, mob, metric) -> dict
def get_specificity_score(row, segment, cohort, metric, mob) -> float

# 8. RATE CALCULATION FUNCTIONS
def fn_cohort_avg(curves_df, segment, cohort, mob, metric_col, lookback) -> float
def fn_cohort_trend(curves_df, segment, cohort, mob, metric_col) -> float
def fn_donor_cohort(curves_df, segment, donor_cohort, mob, metric_col) -> float
def fn_seg_median(curves_df, segment, mob, metric_col) -> float

# 9. RATE APPLICATION
def apply_approach(curves_df, segment, cohort, mob, metric, methodology) -> dict
def apply_rate_cap(rate, metric, approach) -> float

# 10. RATE LOOKUP BUILDER
def build_rate_lookup(seed, curves, methodology, max_months) -> pd.DataFrame
def build_impairment_lookup(seed, impairment_curves, methodology, max_months) -> pd.DataFrame

# 11. FORECAST ENGINE
def run_one_step(seed_table, rate_lookup, impairment_lookup) -> tuple
def run_forecast(seed, rate_lookup, impairment_lookup, max_months) -> pd.DataFrame

# 12. OUTPUT GENERATION
def generate_summary_output(forecast: pd.DataFrame) -> pd.DataFrame
def generate_details_output(forecast: pd.DataFrame) -> pd.DataFrame
def generate_impairment_output(forecast: pd.DataFrame) -> pd.DataFrame
def generate_validation_output(forecast: pd.DataFrame) -> pd.DataFrame
def export_to_excel(summary, details, impairment, validation, output_dir)

# 13. MAIN ORCHESTRATION
def run_backbook_forecast(fact_raw_path, methodology_path, debt_sale_path, output_dir, max_months)
```

---

## Part 3: Detailed Function Specifications

### 3.1 Data Loading Functions

#### load_fact_raw(filepath: str) -> pd.DataFrame

**Purpose:** Load and validate historical data

**Input:** Path to Fact_Raw_Full.csv

**Processing:**
1. Read CSV file
2. Parse CalendarMonth to datetime (handle M/D/YYYY and MM/DD/YYYY)
3. Convert Cohort to string
4. Ensure all numeric columns are float
5. Fill NaN values with 0 for numeric columns
6. Validate required columns exist
7. Sort by CalendarMonth, Segment, Cohort, MOB

**Output:** DataFrame with columns:
- CalendarMonth (datetime)
- Cohort (string)
- Segment (string)
- MOB (int)
- OpeningGBV, NewLoanAmount, Coll_Principal, Coll_Interest, InterestRevenue
- WO_DebtSold, WO_Other, ContraSettlements_Principal, ContraSettlements_Interest
- ClosingGBV_Reported, DaysInMonth
- Provision_Balance, Debt_Sale_WriteOffs, Debt_Sale_Provision_Release, Debt_Sale_Proceeds

**Error Handling:**
- Raise FileNotFoundError if file doesn't exist
- Raise ValueError if required columns missing
- Log warnings for NaN values

---

#### load_rate_methodology(filepath: str) -> pd.DataFrame

**Purpose:** Load rate calculation control table

**Input:** Path to Rate_Methodology.csv

**Processing:**
1. Read CSV file
2. Fill NaN Segment, Cohort, Metric with "ALL"
3. Convert Cohort to string (remove .0 suffix)
4. Ensure MOB_Start, MOB_End are integers
5. Clean Approach and Param1, Param2 strings
6. Validate Approach values are recognized

**Output:** DataFrame with columns:
- Segment (string)
- Cohort (string)
- Metric (string)
- MOB_Start (int)
- MOB_End (int)
- Approach (string)
- Param1 (string or None)
- Param2 (string or None)

---

#### load_debt_sale_schedule(filepath: str) -> pd.DataFrame

**Purpose:** Load debt sale assumptions (optional)

**Input:** Path to Debt_Sale_Schedule.csv

**Processing:**
1. Read CSV file
2. Parse ForecastMonth to datetime
3. Convert Cohort to string
4. Ensure numeric columns are float
5. Sort by ForecastMonth, Segment, Cohort

**Output:** DataFrame with columns:
- ForecastMonth (datetime)
- Segment (string)
- Cohort (string)
- Debt_Sale_WriteOffs (float)
- Debt_Sale_Coverage_Ratio (float)
- Debt_Sale_Proceeds_Rate (float)

---

### 3.2 Curves Calculation Functions

#### calculate_curves_base(fact_raw: pd.DataFrame) -> pd.DataFrame

**Purpose:** Calculate historical rates from actuals

**Processing:**
1. Group by Segment, Cohort, MOB
2. Sum: OpeningGBV, NewLoanAmount, Coll_Principal, Coll_Interest, InterestRevenue, WO_DebtSold, WO_Other, ContraSettlements_Principal, ContraSettlements_Interest
3. Average: DaysInMonth
4. Calculate rates:
   ```
   NewLoanAmount_Rate = NewLoanAmount / OpeningGBV
   Coll_Principal_Rate = Coll_Principal / OpeningGBV
   Coll_Interest_Rate = Coll_Interest / OpeningGBV
   InterestRevenue_Rate = (InterestRevenue / OpeningGBV) * (365 / DaysInMonth)  # Annualized
   WO_DebtSold_Rate = WO_DebtSold / OpeningGBV
   WO_Other_Rate = WO_Other / OpeningGBV
   ContraSettlements_Principal_Rate = ContraSettlements_Principal / OpeningGBV
   ContraSettlements_Interest_Rate = ContraSettlements_Interest / OpeningGBV
   ```
5. Handle division by zero (set rate to 0)
6. Sort by Segment, Cohort, MOB

**Output:** DataFrame with columns:
- Segment, Cohort, MOB
- OpeningGBV, [all amounts]
- [all _Rate columns]

---

#### extend_curves(curves_base: pd.DataFrame, max_months: int) -> pd.DataFrame

**Purpose:** Extend curves beyond max observed MOB for forecasting

**Processing:**
1. For each Segment × Cohort:
   - Find MaxMOB
   - Get rates at MaxMOB
   - For offset in range(1, max_months + 1):
     - Create new row with MOB = MaxMOB + offset
     - Copy all rate columns from MaxMOB row
2. Concatenate with curves_base
3. Sort by Segment, Cohort, MOB

**Output:** Extended curves DataFrame

---

### 3.3 Impairment Curves Functions (NEW)

#### calculate_impairment_actuals(fact_raw: pd.DataFrame) -> pd.DataFrame

**Purpose:** Calculate impairment metrics from historical data

**Processing:**
1. Group by Segment, Cohort, CalendarMonth
2. For each group:
   - Total_Provision_Balance = sum(Provision_Balance)
   - Total_ClosingGBV = sum(ClosingGBV_Reported)
   - Total_Coverage_Ratio = Total_Provision_Balance / Total_ClosingGBV
   - Total_Provision_Movement = Total_Provision_Balance[t] - Total_Provision_Balance[t-1]
   - Debt_Sale_WriteOffs = sum(Debt_Sale_WriteOffs)
   - Debt_Sale_Provision_Release = sum(Debt_Sale_Provision_Release)
   - Debt_Sale_Proceeds = sum(Debt_Sale_Proceeds)
   - If Debt_Sale_WriteOffs > 0:
     - Debt_Sale_Coverage_Ratio = Debt_Sale_Provision_Release / Debt_Sale_WriteOffs
     - Debt_Sale_Proceeds_Rate = Debt_Sale_Proceeds / Debt_Sale_WriteOffs
   - Non_DS_Provision_Movement = Total_Provision_Movement + Debt_Sale_Provision_Release
   - Gross_Impairment_ExcludingDS = Non_DS_Provision_Movement + sum(WO_Other)
   - Debt_Sale_Impact = Debt_Sale_WriteOffs + Debt_Sale_Provision_Release + Debt_Sale_Proceeds
   - Net_Impairment = Gross_Impairment_ExcludingDS + Debt_Sale_Impact

**Output:** DataFrame with columns:
- Segment, Cohort, CalendarMonth
- Total_ClosingGBV, Total_Provision_Balance, Total_Coverage_Ratio
- Total_Provision_Movement
- Debt_Sale_WriteOffs, Debt_Sale_Coverage_Ratio, Debt_Sale_Provision_Release, Debt_Sale_Proceeds_Rate, Debt_Sale_Proceeds
- Non_DS_Provision_Movement, Gross_Impairment_ExcludingDS, Debt_Sale_Impact, Net_Impairment

---

#### calculate_impairment_curves(impairment_actuals: pd.DataFrame) -> pd.DataFrame

**Purpose:** Calculate impairment rates for forecasting

**Processing:**
1. Group by Segment, Cohort, MOB (derived from CalendarMonth)
2. Calculate rates:
   ```
   Total_Coverage_Ratio = Total_Provision_Balance / Total_ClosingGBV
   Debt_Sale_Coverage_Ratio = Debt_Sale_Provision_Release / Debt_Sale_WriteOffs (where DS_WO > 0)
   Debt_Sale_Proceeds_Rate = Debt_Sale_Proceeds / Debt_Sale_WriteOffs (where DS_WO > 0)
   WO_Other_Rate = sum(WO_Other) / Total_ClosingGBV
   ```
3. Handle division by zero
4. Sort by Segment, Cohort, MOB

**Output:** DataFrame with impairment rate columns

---

### 3.4 Seed Generation Functions

#### generate_seed_curves(fact_raw: pd.DataFrame) -> pd.DataFrame

**Purpose:** Create forecast starting point from last month of actuals

**Processing:**
1. Get max CalendarMonth
2. Filter to last month
3. Group by Segment, Cohort:
   - Sum ClosingGBV_Reported → BoM
   - Max MOB → MOB (then add 1)
4. Calculate ForecastMonth = max_cal + 1 month
5. Filter where BoM > 0
6. Select columns: Segment, Cohort, MOB, BoM, ForecastMonth

**Output:** Seed DataFrame with 1 row per Segment × Cohort

---

#### generate_impairment_seed(fact_raw: pd.DataFrame) -> pd.DataFrame

**Purpose:** Create impairment starting point

**Processing:**
1. Get max CalendarMonth
2. Filter to last month
3. Group by Segment, Cohort:
   - Sum Provision_Balance → Prior_Provision_Balance
   - Sum ClosingGBV_Reported → ClosingGBV
4. Calculate ForecastMonth = max_cal + 1 month
5. Select columns: Segment, Cohort, ForecastMonth, Prior_Provision_Balance, ClosingGBV

**Output:** Impairment seed DataFrame

---

### 3.5 Methodology Lookup Functions

#### get_methodology(methodology_df, segment, cohort, mob, metric) -> dict

**Purpose:** Find best matching rate calculation rule

**Processing:**
1. Filter methodology_df where:
   - (Segment == segment OR Segment == "ALL")
   - (Cohort == str(cohort) OR Cohort == "ALL")
   - (Metric == metric OR Metric == "ALL")
   - MOB_Start <= mob <= MOB_End
2. If no matches: return {'Approach': 'NoMatch_ERROR', 'Param1': None, 'Param2': None}
3. Calculate specificity score for each match:
   ```
   score = (segment_match ? 8 : 0) 
         + (cohort_match ? 4 : 0) 
         + (metric_match ? 2 : 0) 
         + 1 / (1 + MOB_End - MOB_Start)
   ```
4. Return row with highest score
5. Extract Approach, Param1, Param2

**Output:** Dictionary with keys:
- Approach (string)
- Param1 (string or None)
- Param2 (string or None)

---

### 3.6 Rate Calculation Functions

#### fn_cohort_avg(curves_df, segment, cohort, mob, metric_col, lookback) -> float

**Purpose:** Calculate average rate from last N MOBs

**Processing:**
1. Filter curves_df where:
   - Segment == segment
   - Cohort == str(cohort)
   - MOB > 3 (post-MOB 3)
   - MOB <= mob
2. Sort by MOB descending
3. Take first lookback rows
4. Return mean of metric_col
5. Return None if < 2 data points

---

#### fn_cohort_trend(curves_df, segment, cohort, mob, metric_col) -> float

**Purpose:** Linear regression extrapolation

**Processing:**
1. Filter curves_df where:
   - Segment == segment
   - Cohort == str(cohort)
   - MOB > 3
   - MOB < mob
2. If < 2 data points: return None
3. Linear regression: y = a + b*x where x=MOB, y=metric_col
4. Calculate: a = mean(y) - b*mean(x)
5. Predict: y[mob] = a + b*mob
6. Return predicted rate

---

#### fn_donor_cohort(curves_df, segment, donor_cohort, mob, metric_col) -> float

**Purpose:** Copy rate from donor cohort

**Processing:**
1. Clean donor_cohort (remove .0 suffix)
2. Filter curves_df where:
   - Segment == segment
   - Cohort == donor_cohort
   - MOB == mob
3. Return metric_col value
4. Return None if not found

---

#### fn_seg_median(curves_df, segment, mob, metric_col) -> float

**Purpose:** Median rate across all cohorts in segment

**Processing:**
1. Filter curves_df where:
   - Segment == segment
   - MOB == mob
2. Return median of metric_col
3. Return None if no data

---

### 3.7 Rate Application Functions

#### apply_approach(curves_df, segment, cohort, mob, metric, methodology) -> dict

**Purpose:** Calculate rate using specified approach

**Processing:**
```python
approach = methodology['Approach']
param1 = methodology['Param1']
metric_col = f"{metric}_Rate"

if approach == 'NoMatch_ERROR':
    return {'Rate': 0, 'ApproachTag': 'NoMatch_ERROR'}
elif approach == 'Zero':
    return {'Rate': 0, 'ApproachTag': 'Zero'}
elif approach == 'Manual':
    try:
        rate = float(param1)
        return {'Rate': rate, 'ApproachTag': 'Manual'}
    except:
        return {'Rate': None, 'ApproachTag': 'Manual_InvalidParam_ERROR'}
elif approach == 'CohortAvg':
    lookback = int(float(param1)) if param1 else 6
    rate = fn_cohort_avg(curves_df, segment, cohort, mob, metric_col, lookback)
    return {'Rate': rate, 'ApproachTag': 'CohortAvg' if rate else 'CohortAvg_NoData_ERROR'}
elif approach == 'CohortTrend':
    rate = fn_cohort_trend(curves_df, segment, cohort, mob, metric_col)
    return {'Rate': rate, 'ApproachTag': 'CohortTrend' if rate else 'CohortTrend_NoData_ERROR'}
elif approach == 'SegMedian':
    rate = fn_seg_median(curves_df, segment, mob, metric_col)
    return {'Rate': rate, 'ApproachTag': 'SegMedian' if rate else 'SegMedian_NoData_ERROR'}
elif approach == 'DonorCohort':
    donor = str(param1).replace('.0', '')
    rate = fn_donor_cohort(curves_df, segment, donor, mob, metric_col)
    return {'Rate': rate, 'ApproachTag': f'DonorCohort:{donor}' if rate else f'DonorCohort_NoData_ERROR:{donor}'}
else:
    return {'Rate': None, 'ApproachTag': f'UnknownApproach_ERROR:{approach}'}
```

**Output:** Dictionary with Rate (float or None) and ApproachTag (string)

---

#### apply_rate_cap(rate, metric, approach) -> float

**Purpose:** Cap rates to reasonable ranges

**Processing:**
1. If rate is None: return 0
2. If 'ERROR' in approach or approach == 'Manual': return rate (no cap)
3. If metric in Config.RATE_CAPS:
   - min_cap, max_cap = Config.RATE_CAPS[metric]
   - return max(min_cap, min(max_cap, rate))
4. Else: return rate

---

### 3.8 Forecast Engine Functions

#### run_one_step(seed_table, rate_lookup, impairment_lookup) -> tuple

**Purpose:** Execute one month of forecast

**Processing:**
1. Join seed_table with rate_lookup on (Segment, Cohort, MOB)
2. Join with impairment_lookup on (Segment, Cohort, MOB)
3. Calculate amounts:
   ```
   OpeningGBV = BoM
   NewLoanAmount = OpeningGBV * NewLoanAmount_Rate
   Coll_Principal = OpeningGBV * Coll_Principal_Rate
   Coll_Interest = OpeningGBV * Coll_Interest_Rate
   InterestRevenue = OpeningGBV * InterestRevenue_Rate / 12
   WO_DebtSold = OpeningGBV * WO_DebtSold_Rate
   WO_Other = OpeningGBV * WO_Other_Rate
   ContraSettlements_Principal = OpeningGBV * ContraSettlements_Principal_Rate
   ContraSettlements_Interest = OpeningGBV * ContraSettlements_Interest_Rate
   ```
4. Calculate ClosingGBV:
   ```
   ClosingGBV = OpeningGBV + InterestRevenue - Coll_Principal - Coll_Interest - WO_DebtSold - WO_Other
   ```
5. Calculate impairment:
   ```
   Total_Provision_Balance = ClosingGBV * Total_Coverage_Ratio
   Total_Provision_Movement = Total_Provision_Balance - Prior_Provision_Balance
   
   If Debt_Sale_WriteOffs > 0:
       Debt_Sale_Provision_Release = Debt_Sale_Coverage_Ratio * Debt_Sale_WriteOffs
       Debt_Sale_Proceeds = Debt_Sale_Proceeds_Rate * Debt_Sale_WriteOffs
   Else:
       Debt_Sale_Provision_Release = 0
       Debt_Sale_Proceeds = 0
   
   Non_DS_Provision_Movement = Total_Provision_Movement + Debt_Sale_Provision_Release
   Gross_Impairment_ExcludingDS = Non_DS_Provision_Movement + WO_Other
   Debt_Sale_Impact = Debt_Sale_WriteOffs + Debt_Sale_Provision_Release + Debt_Sale_Proceeds
   Net_Impairment = Gross_Impairment_ExcludingDS + Debt_Sale_Impact
   ```
6. Calculate ClosingNBV:
   ```
   ClosingNBV = ClosingGBV - Net_Impairment
   ```
7. Prepare output row with all columns
8. Prepare next_seed:
   ```
   next_seed[Segment, Cohort].BoM = ClosingGBV
   next_seed[Segment, Cohort].MOB = MOB + 1
   next_seed[Segment, Cohort].ForecastMonth = ForecastMonth + 1 month
   ```
9. Filter next_seed where BoM > 0

**Output:** Tuple of (step_output_df, next_seed_df)

---

#### run_forecast(seed, rate_lookup, impairment_lookup, max_months) -> pd.DataFrame

**Purpose:** Run complete forecast loop

**Processing:**
1. Initialize all_outputs = []
2. current_seed = seed
3. For month in range(max_months):
   - If len(current_seed) == 0: break
   - step_output, next_seed = run_one_step(current_seed, rate_lookup, impairment_lookup)
   - all_outputs.append(step_output)
   - current_seed = next_seed
4. Concatenate all_outputs
5. Sort by ForecastMonth, Segment, Cohort, MOB
6. Return forecast DataFrame

---

### 3.9 Output Generation Functions

#### generate_summary_output(forecast: pd.DataFrame) -> pd.DataFrame

**Purpose:** Create high-level summary for Excel

**Processing:**
1. Group by ForecastMonth, Segment
2. Sum: OpeningGBV, InterestRevenue, Coll_Principal, Coll_Interest, WO_DebtSold, WO_Other, ClosingGBV, Net_Impairment, ClosingNBV
3. Calculate: Total_Coverage_Ratio = sum(Total_Provision_Balance) / sum(ClosingGBV)
4. Select columns: ForecastMonth, Segment, OpeningGBV, InterestRevenue, Coll_Principal, Coll_Interest, WO_DebtSold, WO_Other, ClosingGBV, Total_Coverage_Ratio, Net_Impairment, ClosingNBV

**Output:** Summary DataFrame

---

#### generate_details_output(forecast: pd.DataFrame) -> pd.DataFrame

**Purpose:** Create complete forecast for Excel

**Processing:**
1. Select all columns from forecast
2. Format dates as YYYY-MM-DD
3. Round numeric columns to appropriate decimals
4. Sort by ForecastMonth, Segment, Cohort, MOB

**Output:** Details DataFrame

---

#### generate_impairment_output(forecast: pd.DataFrame) -> pd.DataFrame

**Purpose:** Create impairment-specific analysis

**Processing:**
1. Select impairment-related columns:
   - ForecastMonth, Segment, Cohort, MOB
   - ClosingGBV, Total_Coverage_Ratio, Total_Provision_Balance, Total_Provision_Movement
   - Debt_Sale_WriteOffs, Debt_Sale_Coverage_Ratio, Debt_Sale_Provision_Release, Debt_Sale_Proceeds
   - Non_DS_Provision_Movement, Gross_Impairment_ExcludingDS, Debt_Sale_Impact, Net_Impairment
2. Sort by ForecastMonth, Segment, Cohort

**Output:** Impairment DataFrame

---

#### generate_validation_output(forecast: pd.DataFrame) -> pd.DataFrame

**Purpose:** Create validation checks

**Processing:**
1. For each row, verify:
   - ClosingGBV = OpeningGBV + InterestRevenue - Coll_Principal - Coll_Interest - WO_DebtSold - WO_Other
   - ClosingNBV = ClosingGBV - Net_Impairment
   - No NaN or infinite values
2. Calculate variance for each check
3. Flag rows with variance > 0.01
4. Create summary of validation issues

**Output:** Validation DataFrame

---

#### export_to_excel(summary, details, impairment, validation, output_dir)

**Purpose:** Write all outputs to Excel workbooks

**Processing:**
1. Create output_dir if not exists
2. Create Forecast_Summary.xlsx:
   - Sheet "Summary" = summary DataFrame
3. Create Forecast_Details.xlsx:
   - Sheet "All_Cohorts" = details DataFrame
4. Create Impairment_Analysis.xlsx:
   - Sheet "Impairment_Detail" = impairment DataFrame
   - Sheet "Coverage_Ratios" = coverage ratio trends
5. Create Validation_Report.xlsx:
   - Sheet "Reconciliation" = validation DataFrame
   - Sheet "Validation_Checks" = summary of issues
6. Format all sheets:
   - Freeze header row
   - Auto-fit columns
   - Format numbers (2-4 decimals as appropriate)
   - Format dates as YYYY-MM-DD

---

### 3.10 Main Orchestration Function

#### run_backbook_forecast(fact_raw_path, methodology_path, debt_sale_path, output_dir, max_months)

**Purpose:** Orchestrate entire forecast process

**Processing:**
```python
# 1. Load data
print("Loading data...")
fact_raw = load_fact_raw(fact_raw_path)
methodology = load_rate_methodology(methodology_path)
debt_sale_schedule = load_debt_sale_schedule(debt_sale_path) if debt_sale_path else None

# 2. Calculate curves
print("Calculating curves...")
curves_base = calculate_curves_base(fact_raw)
curves_extended = extend_curves(curves_base, max_months)

# 3. Calculate impairment curves
print("Calculating impairment curves...")
impairment_actuals = calculate_impairment_actuals(fact_raw)
impairment_curves = calculate_impairment_curves(impairment_actuals)
impairment_curves_extended = extend_curves(impairment_curves, max_months)

# 4. Generate seeds
print("Generating seeds...")
seed = generate_seed_curves(fact_raw)
impairment_seed = generate_impairment_seed(fact_raw)

# 5. Build rate lookups
print("Building rate lookups...")
rate_lookup = build_rate_lookup(seed, curves_extended, methodology, max_months)
impairment_lookup = build_impairment_lookup(seed, impairment_curves_extended, methodology, max_months, debt_sale_schedule)

# 6. Run forecast
print("Running forecast...")
forecast = run_forecast(seed, rate_lookup, impairment_lookup, max_months)

# 7. Generate outputs
print("Generating outputs...")
summary = generate_summary_output(forecast)
details = generate_details_output(forecast)
impairment = generate_impairment_output(forecast)
validation = generate_validation_output(forecast)

# 8. Export to Excel
print("Exporting to Excel...")
export_to_excel(summary, details, impairment, validation, output_dir)

print(f"Forecast complete. Outputs saved to {output_dir}")
```

---

## Part 4: Command Line Interface

```bash
python backbook_forecast.py \
    --fact-raw Fact_Raw_Full.csv \
    --methodology Rate_Methodology.csv \
    --debt-sale Debt_Sale_Schedule.csv \
    --output output/ \
    --months 12
```

**Arguments:**
- `--fact-raw` (required): Path to Fact_Raw_Full.csv
- `--methodology` (required): Path to Rate_Methodology.csv
- `--debt-sale` (optional): Path to Debt_Sale_Schedule.csv
- `--output` (optional): Output directory (default: output/)
- `--months` (optional): Forecast months (default: 12)

---

## Part 5: Error Handling & Logging

### 5.1 Error Handling Strategy

```python
try:
    # Load data
except FileNotFoundError as e:
    print(f"ERROR: File not found: {e}")
    sys.exit(1)
except ValueError as e:
    print(f"ERROR: Invalid data format: {e}")
    sys.exit(1)
except Exception as e:
    print(f"ERROR: Unexpected error: {e}")
    import traceback
    traceback.print_exc()
    sys.exit(1)
```

### 5.2 Logging

```python
import logging

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

logger.info("Loading data...")
logger.warning("Missing data for cohort 202001")
logger.error("Invalid rate value: -0.5")
```

---

## Part 6: Testing & Validation

### 6.1 Unit Tests

```python
def test_load_fact_raw():
    df = load_fact_raw('Fact_Raw_Full.csv')
    assert len(df) > 0
    assert 'CalendarMonth' in df.columns
    assert df['CalendarMonth'].dtype == 'datetime64[ns]'

def test_calculate_curves_base():
    df = load_fact_raw('Fact_Raw_Full.csv')
    curves = calculate_curves_base(df)
    assert 'Coll_Principal_Rate' in curves.columns
    assert curves['Coll_Principal_Rate'].min() >= -1

def test_fn_cohort_avg():
    df = load_fact_raw('Fact_Raw_Full.csv')
    curves = calculate_curves_base(df)
    rate = fn_cohort_avg(curves, 'NRP-S', '202001', 60, 'Coll_Principal_Rate', 6)
    assert rate is not None
    assert -1 <= rate <= 0
```

### 6.2 Integration Tests

```python
def test_full_forecast():
    forecast = run_backbook_forecast(
        'Fact_Raw_Full.csv',
        'Rate_Methodology.csv',
        None,
        'output/',
        3
    )
    assert len(forecast) > 0
    assert 'ClosingGBV' in forecast.columns
    assert 'ClosingNBV' in forecast.columns
    # Verify ClosingGBV = OpeningGBV + InterestRevenue - Collections - WriteOffs
    variance = (forecast['ClosingGBV'] - (forecast['OpeningGBV'] + forecast['InterestRevenue'] - forecast['Coll_Principal'] - forecast['Coll_Interest'] - forecast['WO_DebtSold'] - forecast['WO_Other'])).abs().max()
    assert variance < 0.01
```

---

## Part 7: Dependencies & Requirements

```
# requirements.txt
pandas>=2.0.0
numpy>=1.24.0
openpyxl>=3.10.0
python-dateutil>=2.8.2
```

---

## Part 8: Code Quality Standards

### 8.1 Style Guide
- Follow PEP 8
- Use type hints for all functions
- Use docstrings for all functions and classes
- Max line length: 100 characters
- Use meaningful variable names

### 8.2 Documentation
- Include docstring for every function
- Include examples in docstrings
- Comment complex logic
- Include error handling documentation

### 8.3 Performance
- Use vectorized pandas operations (avoid loops)
- Cache intermediate results
- Use efficient data types (float32 where appropriate)
- Profile code for bottlenecks

---

## Part 9: Deliverables Checklist

- [ ] backbook_forecast.py (single module with all code)
- [ ] requirements.txt (dependencies)
- [ ] README.md (usage instructions)
- [ ] Sample Fact_Raw_Full.csv
- [ ] Sample Rate_Methodology.csv
- [ ] Sample Debt_Sale_Schedule.csv
- [ ] Output Excel files (4 workbooks)
- [ ] Unit tests
- [ ] Integration tests
- [ ] Error handling & logging
- [ ] Command line interface

---

## Part 10: Known Limitations & Future Enhancements

### 10.1 Known Limitations
- Assumes linear trend extrapolation (CohortTrend)
- Debt sale schedule must be provided separately
- No seasonality adjustment
- No macroeconomic adjustments

### 10.2 Future Enhancements
- Add polynomial trend extrapolation
- Add seasonality factors
- Add macroeconomic adjustments
- Add Monte Carlo simulation for sensitivity analysis
- Add interactive dashboard (Streamlit)
- Add database backend for large datasets
