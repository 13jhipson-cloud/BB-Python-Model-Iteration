#!/usr/bin/env python3
"""
Forecast Calibrator Tool
========================
A comprehensive tool for iteratively calibrating forecast outputs to match budget targets.

This tool:
1. Compares forecast outputs to budget consol targets
2. Tests alternative approaches for each rate metric
3. Recommends approach changes that minimize variance
4. Documents business rationale for each change
5. Identifies remaining gaps for overlay consideration

Key principles:
- Uses ONLY existing forecast approaches (CohortAvg, CohortTrend, DonorCohort, etc.)
- Never artificially scales numbers
- All changes must be justifiable with business logic
- Accounts for cascade effects between metrics

Author: Forecast Calibration System
"""

import pandas as pd
import numpy as np
import logging
import argparse
import os
import sys
import subprocess
import tempfile
import shutil
from datetime import datetime
from typing import Dict, List, Tuple, Optional, Any
from dataclasses import dataclass
from enum import Enum
import warnings
warnings.filterwarnings('ignore')

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)


# =============================================================================
# SECTION 1: CONFIGURATION AND CONSTANTS
# =============================================================================

class Config:
    """Global configuration for the calibrator."""

    # File paths (can be overridden via command line)
    BUDGET_FILE = "Budget consol file.xlsx"
    FORECAST_SCRIPT = "backbook_forecast.py"
    METHODOLOGY_FILE = "Rate_Methodology.csv"
    FACT_RAW_FILE = "Fact_Raw_New.xlsx"

    # Budget file structure (row indices in P&L analysis - BB sheet)
    # Row 2 contains dates (columns 4 onwards)
    BUDGET_DATE_ROW = 2
    BUDGET_DATA_START_COL = 4

    # Budget row mappings (0-indexed)
    BUDGET_ROWS = {
        'Collections': {
            'Non Prime': 11,
            'Near Prime Small': 12,
            'Near Prime Medium': 13,
            'Prime': 14,
            'Total': 15
        },
        'ClosingGBV': {
            'Non Prime': 22,
            'Near Prime Small': 23,
            'Near Prime Medium': 24,
            'Prime': 25,
            'Total': 26
        },
        'ClosingNBV': {
            'Non Prime': 42,
            'Near Prime Small': 43,
            'Near Prime Medium': 44,
            'Prime': 45,
            'Total': 46
        },
        'Revenue': {
            'Non Prime': 62,
            'Near Prime Small': 63,
            'Near Prime Medium': 64,
            'Prime': 65,
            'Total': 66
        },
        'GrossImpairment': {
            'Non Prime': 73,
            'Near Prime Small': 74,
            'Near Prime Medium': 75,
            'Prime': 76,
            'Total': 77
        },
        'NetImpairment': {
            'Non Prime': 121,
            'Near Prime Small': 122,
            'Near Prime Medium': 123,
            'Prime': 124,
            'Total': 125
        }
    }

    # Segment mapping: Budget name -> Forecast name(s)
    SEGMENT_MAPPING = {
        'Non Prime': ['NON PRIME'],
        'Near Prime Small': ['NRP-S'],
        'Near Prime Medium': ['NRP-M', 'NRP-L'],  # NRP-L grouped with NRP-M
        'Prime': ['PRIME']
    }

    # Reverse mapping: Forecast name -> Budget name
    FORECAST_TO_BUDGET_SEGMENT = {
        'NON PRIME': 'Non Prime',
        'NRP-S': 'Near Prime Small',
        'NRP-M': 'Near Prime Medium',
        'NRP-L': 'Near Prime Medium',  # NRP-L maps to Near Prime Medium
        'PRIME': 'Prime'
    }

    # Available approaches for testing
    AVAILABLE_APPROACHES = [
        'CohortAvg',
        'CohortTrend',
        'DonorCohort',
        'ScaledDonor',
        'SegMedian',
        'Manual',
        'Zero'
    ]

    # Rate metrics that can be adjusted
    ADJUSTABLE_RATES = [
        'Coll_Principal_Rate',
        'Coll_Interest_Rate',
        'InterestRevenue_Rate',
        'WO_DebtSold_Rate',
        'WO_Other_Rate',
        'Total_Coverage_Ratio'
    ]

    # Lookback periods to test for averaging approaches
    LOOKBACK_PERIODS = [3, 6, 9, 12]

    # Potential donor cohorts (mature cohorts with good data)
    DONOR_COHORTS = ['201912', '202001', '202101', '202201', '202301']

    # Tolerance for considering a variance acceptable (in millions)
    VARIANCE_TOLERANCE_M = 0.5  # £0.5M

    # Maximum iterations for optimization loop
    MAX_ITERATIONS = 50


@dataclass
class BudgetData:
    """Container for budget data by segment and month."""
    collections: Dict[str, Dict[str, float]]  # segment -> month -> value
    closing_gbv: Dict[str, Dict[str, float]]
    closing_nbv: Dict[str, Dict[str, float]]
    revenue: Dict[str, Dict[str, float]]
    gross_impairment: Dict[str, Dict[str, float]]
    net_impairment: Dict[str, Dict[str, float]]
    months: List[str]  # List of forecast months


@dataclass
class ForecastData:
    """Container for forecast data by segment and month."""
    collections: Dict[str, Dict[str, float]]
    closing_gbv: Dict[str, Dict[str, float]]
    closing_nbv: Dict[str, Dict[str, float]]
    revenue: Dict[str, Dict[str, float]]
    gross_impairment: Dict[str, Dict[str, float]]
    net_impairment: Dict[str, Dict[str, float]]
    months: List[str]


@dataclass
class VarianceReport:
    """Container for variance analysis results."""
    metric: str
    segment: str
    month: str
    budget_value: float
    forecast_value: float
    variance: float
    variance_pct: float
    within_tolerance: bool


@dataclass
class ApproachRecommendation:
    """Container for a recommended approach change."""
    segment: str
    cohort: str
    metric: str
    current_approach: str
    current_param: Optional[str]
    recommended_approach: str
    recommended_param: Optional[str]
    expected_impact: Dict[str, float]  # metric -> impact
    rationale: str
    confidence: float  # 0-1 confidence score


# =============================================================================
# SECTION 2: BUDGET DATA LOADING
# =============================================================================

class BudgetLoader:
    """Load and parse budget data from Budget Consol file."""

    def __init__(self, filepath: str):
        self.filepath = filepath
        self.raw_data = None

    def load(self) -> BudgetData:
        """Load budget data from Excel file."""
        logger.info(f"Loading budget data from {self.filepath}")

        try:
            self.raw_data = pd.read_excel(
                self.filepath,
                sheet_name='P&L analysis - BB',
                header=None
            )
        except Exception as e:
            logger.error(f"Failed to load budget file: {e}")
            raise

        # Extract dates from row 2
        months = self._extract_months()
        logger.info(f"Found {len(months)} forecast months in budget: {months[0]} to {months[-1]}")

        # Extract data for each metric
        collections = self._extract_metric('Collections', months)
        closing_gbv = self._extract_metric('ClosingGBV', months)
        closing_nbv = self._extract_metric('ClosingNBV', months)
        revenue = self._extract_metric('Revenue', months)
        gross_impairment = self._extract_metric('GrossImpairment', months)
        net_impairment = self._extract_metric('NetImpairment', months)

        return BudgetData(
            collections=collections,
            closing_gbv=closing_gbv,
            closing_nbv=closing_nbv,
            revenue=revenue,
            gross_impairment=gross_impairment,
            net_impairment=net_impairment,
            months=months
        )

    def _extract_months(self) -> List[str]:
        """Extract month labels from the budget file."""
        months = []
        row = Config.BUDGET_DATE_ROW

        for col in range(Config.BUDGET_DATA_START_COL, self.raw_data.shape[1]):
            val = self.raw_data.iloc[row, col]
            if pd.isna(val):
                break

            # Convert to string date format YYYY-MM
            if isinstance(val, (pd.Timestamp, datetime)):
                month_str = val.strftime('%Y-%m')
            else:
                # Try to parse as date
                try:
                    dt = pd.to_datetime(val)
                    month_str = dt.strftime('%Y-%m')
                except:
                    continue

            months.append(month_str)

        return months

    def _extract_metric(self, metric_name: str, months: List[str]) -> Dict[str, Dict[str, float]]:
        """Extract data for a specific metric across all segments."""
        result = {}

        if metric_name not in Config.BUDGET_ROWS:
            logger.warning(f"Unknown metric: {metric_name}")
            return result

        row_mapping = Config.BUDGET_ROWS[metric_name]

        for segment_name, row_idx in row_mapping.items():
            if segment_name == 'Total':
                continue  # Skip totals, we'll calculate from segments

            segment_data = {}
            for i, month in enumerate(months):
                col = Config.BUDGET_DATA_START_COL + i
                val = self.raw_data.iloc[row_idx, col]

                if pd.notna(val):
                    try:
                        segment_data[month] = float(val)
                    except (ValueError, TypeError):
                        segment_data[month] = 0.0
                else:
                    segment_data[month] = 0.0

            result[segment_name] = segment_data

        return result


# =============================================================================
# SECTION 3: FORECAST DATA LOADING
# =============================================================================

class ForecastLoader:
    """Load and aggregate forecast output data."""

    def __init__(self, forecast_output_path: str = None):
        self.forecast_output_path = forecast_output_path

    def load_from_combined_view(self, filepath: str) -> ForecastData:
        """Load forecast data from a transparency report's Combined View sheet."""
        logger.info(f"Loading forecast data from {filepath}")

        try:
            df = pd.read_excel(filepath, sheet_name='6_Combined_View')
        except Exception as e:
            logger.error(f"Failed to load forecast file: {e}")
            raise

        # Filter to forecast data only
        forecast_df = df[df['DataType'] == 'Forecast'].copy()

        if len(forecast_df) == 0:
            logger.warning("No forecast data found in Combined View")
            return None

        # Convert Month to string format
        forecast_df['MonthStr'] = pd.to_datetime(forecast_df['Month']).dt.strftime('%Y-%m')

        # Get unique months
        months = sorted(forecast_df['MonthStr'].unique())

        # Aggregate by segment and month
        collections = self._aggregate_metric(forecast_df, 'Coll_Principal', months, include_interest=True)
        closing_gbv = self._aggregate_metric(forecast_df, 'ClosingGBV', months)
        revenue = self._aggregate_metric(forecast_df, 'InterestRevenue', months)

        # Load impairment data from dedicated sheet if available
        try:
            impairment_df = pd.read_excel(filepath, sheet_name='11_Impairment')
            impairment_df['MonthStr'] = pd.to_datetime(impairment_df['ForecastMonth']).dt.strftime('%Y-%m')
            closing_nbv = self._aggregate_impairment_metric(impairment_df, 'ClosingNBV', months)
            net_impairment = self._aggregate_impairment_metric(impairment_df, 'Net_Impairment', months)
            gross_impairment = self._aggregate_impairment_metric(impairment_df, 'Gross_Impairment_ExcludingDS', months)
        except Exception as e:
            logger.warning(f"Could not load impairment sheet, calculating from Combined View: {e}")
            closing_nbv = self._calculate_nbv(forecast_df, months)
            gross_impairment = self._calculate_impairment(forecast_df, months)
            net_impairment = gross_impairment

        return ForecastData(
            collections=collections,
            closing_gbv=closing_gbv,
            closing_nbv=closing_nbv,
            revenue=revenue,
            gross_impairment=gross_impairment,
            net_impairment=net_impairment,
            months=months
        )

    def _aggregate_impairment_metric(self, df: pd.DataFrame, metric: str,
                                     months: List[str]) -> Dict[str, Dict[str, float]]:
        """Aggregate impairment metric by budget segment and month."""
        result = {}

        for budget_segment, forecast_segments in Config.SEGMENT_MAPPING.items():
            segment_data = {}

            for month in months:
                mask = (df['MonthStr'] == month) & (df['Segment'].isin(forecast_segments))
                month_df = df[mask]

                if len(month_df) == 0:
                    segment_data[month] = 0.0
                    continue

                if metric in month_df.columns:
                    segment_data[month] = month_df[metric].sum()
                else:
                    segment_data[month] = 0.0

            result[budget_segment] = segment_data

        return result

    def _aggregate_metric(self, df: pd.DataFrame, metric: str, months: List[str],
                         include_interest: bool = False) -> Dict[str, Dict[str, float]]:
        """Aggregate a metric by budget segment and month."""
        result = {}

        for budget_segment, forecast_segments in Config.SEGMENT_MAPPING.items():
            segment_data = {}

            for month in months:
                # Filter to this month and these segments
                mask = (df['MonthStr'] == month) & (df['Segment'].isin(forecast_segments))
                month_df = df[mask]

                if len(month_df) == 0:
                    segment_data[month] = 0.0
                    continue

                if metric == 'Coll_Principal' and include_interest:
                    # Collections = Principal + Interest
                    principal = month_df['Coll_Principal'].sum() if 'Coll_Principal' in month_df.columns else 0
                    interest = month_df['Coll_Interest'].sum() if 'Coll_Interest' in month_df.columns else 0
                    segment_data[month] = abs(principal) + abs(interest)  # Collections are negative in raw data
                else:
                    if metric in month_df.columns:
                        segment_data[month] = month_df[metric].sum()
                    else:
                        segment_data[month] = 0.0

            result[budget_segment] = segment_data

        return result

    def _calculate_nbv(self, df: pd.DataFrame, months: List[str]) -> Dict[str, Dict[str, float]]:
        """Calculate Closing NBV = Closing GBV - Provision Balance."""
        result = {}

        for budget_segment, forecast_segments in Config.SEGMENT_MAPPING.items():
            segment_data = {}

            for month in months:
                mask = (df['MonthStr'] == month) & (df['Segment'].isin(forecast_segments))
                month_df = df[mask]

                if len(month_df) == 0:
                    segment_data[month] = 0.0
                    continue

                closing_gbv = month_df['ClosingGBV'].sum() if 'ClosingGBV' in month_df.columns else 0
                provision = month_df['Provision_Balance'].sum() if 'Provision_Balance' in month_df.columns else 0

                segment_data[month] = closing_gbv - provision

            result[budget_segment] = segment_data

        return result

    def _calculate_impairment(self, df: pd.DataFrame, months: List[str]) -> Dict[str, Dict[str, float]]:
        """Calculate Net Impairment from provision movements and write-offs."""
        # This is a simplified calculation - actual implementation may need more detail
        result = {}

        for budget_segment, forecast_segments in Config.SEGMENT_MAPPING.items():
            segment_data = {}

            for month in months:
                mask = (df['MonthStr'] == month) & (df['Segment'].isin(forecast_segments))
                month_df = df[mask]

                if len(month_df) == 0:
                    segment_data[month] = 0.0
                    continue

                # Simplified: use WO_Other as proxy for impairment charge
                # In reality, this should be Provision Movement + Write-offs
                wo_other = month_df['WO_Other'].sum() if 'WO_Other' in month_df.columns else 0
                segment_data[month] = wo_other

            result[budget_segment] = segment_data

        return result


# =============================================================================
# SECTION 4: VARIANCE CALCULATION
# =============================================================================

class VarianceCalculator:
    """Calculate and analyze variances between forecast and budget."""

    def __init__(self, budget: BudgetData, forecast: ForecastData):
        self.budget = budget
        self.forecast = forecast

    def calculate_all_variances(self) -> List[VarianceReport]:
        """Calculate variances for all metrics, segments, and months."""
        variances = []

        # Define metric pairs (budget attr, forecast attr)
        metrics = [
            ('Collections', 'collections', 'collections'),
            ('ClosingGBV', 'closing_gbv', 'closing_gbv'),
            ('ClosingNBV', 'closing_nbv', 'closing_nbv'),
            ('Revenue', 'revenue', 'revenue'),
            ('NetImpairment', 'net_impairment', 'net_impairment'),
        ]

        for metric_name, budget_attr, forecast_attr in metrics:
            budget_data = getattr(self.budget, budget_attr)
            forecast_data = getattr(self.forecast, forecast_attr)

            for segment in budget_data.keys():
                if segment not in forecast_data:
                    continue

                for month in self.budget.months:
                    if month not in budget_data[segment] or month not in forecast_data[segment]:
                        continue

                    budget_val = budget_data[segment][month]
                    forecast_val = forecast_data[segment][month]
                    variance = forecast_val - budget_val

                    # Calculate percentage variance (avoid division by zero)
                    if abs(budget_val) > 1:
                        variance_pct = (variance / abs(budget_val)) * 100
                    else:
                        variance_pct = 0 if abs(variance) < 1 else float('inf')

                    # Check if within tolerance
                    within_tolerance = abs(variance) <= (Config.VARIANCE_TOLERANCE_M * 1_000_000)

                    variances.append(VarianceReport(
                        metric=metric_name,
                        segment=segment,
                        month=month,
                        budget_value=budget_val,
                        forecast_value=forecast_val,
                        variance=variance,
                        variance_pct=variance_pct,
                        within_tolerance=within_tolerance
                    ))

        return variances

    def get_summary_by_metric(self, variances: List[VarianceReport]) -> Dict[str, Dict[str, float]]:
        """Summarize total variance by metric across all segments and months."""
        summary = {}

        for v in variances:
            if v.metric not in summary:
                summary[v.metric] = {
                    'total_budget': 0,
                    'total_forecast': 0,
                    'total_variance': 0,
                    'count_outside_tolerance': 0,
                    'max_variance': 0
                }

            summary[v.metric]['total_budget'] += v.budget_value
            summary[v.metric]['total_forecast'] += v.forecast_value
            summary[v.metric]['total_variance'] += v.variance

            if not v.within_tolerance:
                summary[v.metric]['count_outside_tolerance'] += 1

            if abs(v.variance) > abs(summary[v.metric]['max_variance']):
                summary[v.metric]['max_variance'] = v.variance

        return summary

    def get_worst_variances(self, variances: List[VarianceReport], n: int = 10) -> List[VarianceReport]:
        """Get the N worst variances by absolute value."""
        sorted_variances = sorted(variances, key=lambda v: abs(v.variance), reverse=True)
        return sorted_variances[:n]

    def get_variances_by_segment(self, variances: List[VarianceReport]) -> Dict[str, List[VarianceReport]]:
        """Group variances by segment."""
        by_segment = {}
        for v in variances:
            if v.segment not in by_segment:
                by_segment[v.segment] = []
            by_segment[v.segment].append(v)
        return by_segment


# =============================================================================
# SECTION 5: METHODOLOGY MANAGER
# =============================================================================

class MethodologyManager:
    """Manage rate methodology configurations."""

    def __init__(self, filepath: str):
        self.filepath = filepath
        self.methodology = None
        self.original_methodology = None

    def load(self) -> pd.DataFrame:
        """Load methodology from CSV file."""
        logger.info(f"Loading methodology from {self.filepath}")
        self.methodology = pd.read_csv(self.filepath)
        self.original_methodology = self.methodology.copy()
        return self.methodology

    def save(self, filepath: str = None):
        """Save methodology to CSV file."""
        if filepath is None:
            filepath = self.filepath
        self.methodology.to_csv(filepath, index=False)
        logger.info(f"Saved methodology to {filepath}")

    def get_current_approach(self, segment: str, cohort: str, metric: str, mob: int = 0) -> Tuple[str, Optional[str]]:
        """Get the current approach for a segment/cohort/metric combination."""
        if self.methodology is None:
            return None, None

        # Filter for matching rules
        mask = (
            (self.methodology['Segment'] == segment) &
            (self.methodology['Metric'] == metric) &
            (self.methodology['MOB_Start'] <= mob) &
            (self.methodology['MOB_End'] >= mob)
        )

        # Check cohort-specific rules first
        cohort_mask = mask & (self.methodology['Cohort'] == cohort)
        if cohort_mask.any():
            rule = self.methodology[cohort_mask].iloc[0]
            return rule['Approach'], str(rule.get('Param1', '')) if pd.notna(rule.get('Param1')) else None

        # Fall back to ALL cohorts
        all_mask = mask & (self.methodology['Cohort'] == 'ALL')
        if all_mask.any():
            rule = self.methodology[all_mask].iloc[0]
            return rule['Approach'], str(rule.get('Param1', '')) if pd.notna(rule.get('Param1')) else None

        return None, None

    def update_approach(self, segment: str, cohort: str, metric: str,
                       new_approach: str, new_param: Optional[str] = None,
                       mob_start: int = 0, mob_end: int = 999):
        """Update or add an approach rule."""
        if self.methodology is None:
            raise ValueError("Methodology not loaded")

        # Check if rule exists
        mask = (
            (self.methodology['Segment'] == segment) &
            (self.methodology['Cohort'] == cohort) &
            (self.methodology['Metric'] == metric) &
            (self.methodology['MOB_Start'] == mob_start) &
            (self.methodology['MOB_End'] == mob_end)
        )

        if mask.any():
            # Update existing rule
            self.methodology.loc[mask, 'Approach'] = new_approach
            if new_param is not None:
                self.methodology.loc[mask, 'Param1'] = new_param
        else:
            # Add new rule
            new_rule = {
                'Segment': segment,
                'Cohort': cohort,
                'Metric': metric,
                'MOB_Start': mob_start,
                'MOB_End': mob_end,
                'Approach': new_approach,
                'Param1': new_param if new_param else np.nan,
                'Param2': np.nan
            }
            self.methodology = pd.concat([
                self.methodology,
                pd.DataFrame([new_rule])
            ], ignore_index=True)

    def get_all_cohorts(self, segment: str) -> List[str]:
        """Get all cohorts defined for a segment."""
        if self.methodology is None:
            return []

        mask = self.methodology['Segment'] == segment
        cohorts = self.methodology[mask]['Cohort'].unique().tolist()
        return [c for c in cohorts if c != 'ALL']

    def reset_to_original(self):
        """Reset methodology to original loaded state."""
        if self.original_methodology is not None:
            self.methodology = self.original_methodology.copy()


# =============================================================================
# SECTION 6: FORECAST RUNNER
# =============================================================================

class ForecastRunner:
    """Run the forecast model with different configurations."""

    def __init__(self, forecast_script: str, fact_raw_file: str):
        self.forecast_script = forecast_script
        self.fact_raw_file = fact_raw_file

    def run_forecast(self, methodology_file: str, output_dir: str, months: int = 10) -> str:
        """
        Run the forecast with specified methodology.

        Returns path to the generated transparency report.
        """
        logger.info(f"Running forecast with methodology: {methodology_file}")

        cmd = [
            'python3', self.forecast_script,
            '--fact-raw', self.fact_raw_file,
            '--methodology', methodology_file,
            '--months', str(months),
            '--output', output_dir,
            '--transparency-report'
        ]

        try:
            result = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                timeout=300  # 5 minute timeout
            )

            if result.returncode != 0:
                logger.error(f"Forecast failed: {result.stderr}")
                return None

            # Find the transparency report in output directory
            report_path = os.path.join(output_dir, 'Forecast_Transparency_Report.xlsx')
            if os.path.exists(report_path):
                return report_path

            logger.error(f"Transparency report not found at {report_path}")
            return None

        except subprocess.TimeoutExpired:
            logger.error("Forecast timed out after 5 minutes")
            return None
        except Exception as e:
            logger.error(f"Error running forecast: {e}")
            return None


# =============================================================================
# SECTION 7: APPROACH TESTER
# =============================================================================

class ApproachTester:
    """Test different approaches and measure their impact."""

    def __init__(self, methodology_manager: MethodologyManager,
                 forecast_runner: ForecastRunner,
                 budget: BudgetData):
        self.methodology_manager = methodology_manager
        self.forecast_runner = forecast_runner
        self.budget = budget
        self.test_results_cache = {}

    def test_approach(self, segment: str, cohort: str, metric: str,
                     approach: str, param: Optional[str] = None) -> Dict[str, float]:
        """
        Test a specific approach and return its impact on all metrics.

        Returns dict of metric -> total variance change
        """
        # Create cache key
        cache_key = f"{segment}|{cohort}|{metric}|{approach}|{param}"
        if cache_key in self.test_results_cache:
            return self.test_results_cache[cache_key]

        # Create temporary methodology file
        temp_dir = tempfile.mkdtemp()
        temp_methodology = os.path.join(temp_dir, 'test_methodology.csv')

        try:
            # Save current methodology
            self.methodology_manager.methodology.to_csv(temp_methodology, index=False)

            # Update with test approach
            test_manager = MethodologyManager(temp_methodology)
            test_manager.load()
            test_manager.update_approach(segment, cohort, metric, approach, param)
            test_manager.save()

            # Run forecast
            output_dir = os.path.join(temp_dir, 'output')
            os.makedirs(output_dir, exist_ok=True)

            report_path = self.forecast_runner.run_forecast(
                temp_methodology, output_dir, months=10
            )

            if report_path is None:
                logger.warning(f"Forecast failed for approach test: {cache_key}")
                return None

            # Load forecast results
            loader = ForecastLoader()
            forecast = loader.load_from_combined_view(report_path)

            if forecast is None:
                return None

            # Calculate variances
            calc = VarianceCalculator(self.budget, forecast)
            variances = calc.calculate_all_variances()
            summary = calc.get_summary_by_metric(variances)

            # Extract total variances
            result = {}
            for metric_name, data in summary.items():
                result[metric_name] = data['total_variance']

            # Cache result
            self.test_results_cache[cache_key] = result
            return result

        finally:
            # Cleanup
            shutil.rmtree(temp_dir, ignore_errors=True)

    def find_best_approach(self, segment: str, cohort: str, metric: str,
                          current_variance: Dict[str, float]) -> Optional[ApproachRecommendation]:
        """
        Find the best approach for a given segment/cohort/metric.

        Tests all available approaches and returns the one that minimizes
        overall variance while staying within constraints.
        """
        best_result = None
        best_improvement = 0
        best_approach = None
        best_param = None

        current_approach, current_param = self.methodology_manager.get_current_approach(
            segment, cohort, metric
        )

        # Test each available approach
        for approach in Config.AVAILABLE_APPROACHES:
            params_to_test = [None]

            # Add specific parameters based on approach type
            if approach in ['CohortAvg', 'CohortTrend']:
                params_to_test = [str(p) for p in Config.LOOKBACK_PERIODS]
            elif approach == 'DonorCohort':
                params_to_test = Config.DONOR_COHORTS

            for param in params_to_test:
                # Skip if same as current
                if approach == current_approach and str(param) == str(current_param):
                    continue

                # Test this approach
                result = self.test_approach(segment, cohort, metric, approach, param)

                if result is None:
                    continue

                # Calculate improvement (reduction in total absolute variance)
                current_total = sum(abs(v) for v in current_variance.values())
                new_total = sum(abs(v) for v in result.values())
                improvement = current_total - new_total

                if improvement > best_improvement:
                    best_improvement = improvement
                    best_result = result
                    best_approach = approach
                    best_param = param

        if best_approach is None:
            return None

        # Generate rationale
        rationale = self._generate_rationale(
            segment, cohort, metric,
            current_approach, current_param,
            best_approach, best_param,
            best_improvement
        )

        return ApproachRecommendation(
            segment=segment,
            cohort=cohort,
            metric=metric,
            current_approach=current_approach,
            current_param=current_param,
            recommended_approach=best_approach,
            recommended_param=best_param,
            expected_impact=best_result,
            rationale=rationale,
            confidence=min(1.0, best_improvement / 1_000_000)  # Normalize by £1M
        )

    def _generate_rationale(self, segment: str, cohort: str, metric: str,
                           current_approach: str, current_param: str,
                           new_approach: str, new_param: str,
                           improvement: float) -> str:
        """Generate business rationale for an approach change."""
        rationale_parts = []

        # Describe the change
        if new_approach == 'DonorCohort':
            rationale_parts.append(
                f"Use cohort {new_param} as donor for {metric} rates. "
                f"Cohort {new_param} has mature, stable rate patterns that align with budget assumptions."
            )
        elif new_approach == 'CohortTrend':
            rationale_parts.append(
                f"Switch to trend-based forecasting with {new_param}-period lookback. "
                f"Recent actuals show a directional trend that should be extrapolated."
            )
        elif new_approach == 'CohortAvg':
            rationale_parts.append(
                f"Use {new_param}-period rolling average. "
                f"Rates have stabilized and historical average provides best estimate."
            )
        elif new_approach == 'SegMedian':
            rationale_parts.append(
                f"Fall back to segment median rates due to insufficient cohort-specific data."
            )

        # Add improvement context
        if improvement > 0:
            improvement_m = improvement / 1_000_000
            rationale_parts.append(
                f"Expected to reduce total variance by £{improvement_m:.2f}M."
            )

        return " ".join(rationale_parts)


# =============================================================================
# SECTION 8: OPTIMIZATION ENGINE
# =============================================================================

class OptimizationEngine:
    """Main optimization engine for calibrating forecast to budget."""

    def __init__(self, budget: BudgetData,
                 methodology_manager: MethodologyManager,
                 forecast_runner: ForecastRunner):
        self.budget = budget
        self.methodology_manager = methodology_manager
        self.forecast_runner = forecast_runner
        self.recommendations = []
        self.iteration_history = []

    def optimize(self, max_iterations: int = None) -> List[ApproachRecommendation]:
        """
        Run the optimization loop to minimize variance.

        Uses tiered optimization:
        1. GBV-affecting rates first (Collections, WriteOffs)
        2. Coverage/Impairment rates second
        3. Revenue rates last
        """
        if max_iterations is None:
            max_iterations = Config.MAX_ITERATIONS

        logger.info("Starting optimization...")

        # Create approach tester
        tester = ApproachTester(
            self.methodology_manager,
            self.forecast_runner,
            self.budget
        )

        iteration = 0
        while iteration < max_iterations:
            iteration += 1
            logger.info(f"=== Optimization Iteration {iteration} ===")

            # Run current forecast
            temp_dir = tempfile.mkdtemp()
            try:
                temp_meth = os.path.join(temp_dir, 'current_methodology.csv')
                self.methodology_manager.save(temp_meth)

                output_dir = os.path.join(temp_dir, 'output')
                os.makedirs(output_dir)

                report_path = self.forecast_runner.run_forecast(
                    temp_meth, output_dir, months=10
                )

                if report_path is None:
                    logger.error("Forecast failed, stopping optimization")
                    break

                # Load and calculate variances
                loader = ForecastLoader()
                forecast = loader.load_from_combined_view(report_path)

                calc = VarianceCalculator(self.budget, forecast)
                variances = calc.calculate_all_variances()
                summary = calc.get_summary_by_metric(variances)

                # Store iteration results
                self.iteration_history.append({
                    'iteration': iteration,
                    'summary': summary.copy(),
                    'total_variance': sum(abs(s['total_variance']) for s in summary.values())
                })

                # Check if we're within tolerance
                all_within_tolerance = all(
                    abs(s['total_variance']) <= Config.VARIANCE_TOLERANCE_M * 1_000_000 * 10  # Per-metric tolerance
                    for s in summary.values()
                )

                if all_within_tolerance:
                    logger.info("All metrics within tolerance. Optimization complete.")
                    break

                # Find worst variance and try to improve
                current_variance = {m: s['total_variance'] for m, s in summary.items()}
                worst_metric = max(current_variance.keys(), key=lambda m: abs(current_variance[m]))

                logger.info(f"Worst variance: {worst_metric} = £{current_variance[worst_metric]/1e6:.2f}M")

                # Try to find improvement for each segment
                found_improvement = False

                for segment in Config.SEGMENT_MAPPING.keys():
                    forecast_segments = Config.SEGMENT_MAPPING[segment]

                    for fs in forecast_segments:
                        # Get cohorts for this segment
                        cohorts = self.methodology_manager.get_all_cohorts(fs)
                        if not cohorts:
                            cohorts = ['ALL']

                        # Determine which rate metric to adjust based on worst variance
                        rate_metrics = self._get_rate_metrics_for_variance(worst_metric)

                        for rate_metric in rate_metrics:
                            for cohort in cohorts:
                                # Find best approach for this combination
                                recommendation = tester.find_best_approach(
                                    fs, cohort, rate_metric, current_variance
                                )

                                if recommendation and recommendation.confidence > 0.1:
                                    logger.info(
                                        f"Found improvement: {fs}/{cohort}/{rate_metric} "
                                        f"-> {recommendation.recommended_approach}"
                                    )

                                    # Apply the change
                                    self.methodology_manager.update_approach(
                                        fs, cohort, rate_metric,
                                        recommendation.recommended_approach,
                                        recommendation.recommended_param
                                    )

                                    self.recommendations.append(recommendation)
                                    found_improvement = True
                                    break

                            if found_improvement:
                                break

                        if found_improvement:
                            break

                    if found_improvement:
                        break

                if not found_improvement:
                    logger.info("No further improvements found. Stopping optimization.")
                    break

            finally:
                shutil.rmtree(temp_dir, ignore_errors=True)

        logger.info(f"Optimization complete after {iteration} iterations")
        logger.info(f"Generated {len(self.recommendations)} recommendations")

        return self.recommendations

    def _get_rate_metrics_for_variance(self, variance_metric: str) -> List[str]:
        """Map a variance metric to the rate metrics that affect it."""
        mapping = {
            'Collections': ['Coll_Principal_Rate', 'Coll_Interest_Rate'],
            'ClosingGBV': ['Coll_Principal_Rate', 'Coll_Interest_Rate', 'WO_DebtSold_Rate', 'WO_Other_Rate'],
            'ClosingNBV': ['Total_Coverage_Ratio', 'Coll_Principal_Rate'],
            'Revenue': ['InterestRevenue_Rate'],
            'NetImpairment': ['Total_Coverage_Ratio', 'WO_Other_Rate'],
            'GrossImpairment': ['Total_Coverage_Ratio', 'WO_Other_Rate']
        }
        return mapping.get(variance_metric, Config.ADJUSTABLE_RATES)


# =============================================================================
# SECTION 9: REPORT GENERATOR
# =============================================================================

class ReportGenerator:
    """Generate calibration reports and outputs."""

    def __init__(self, budget: BudgetData, recommendations: List[ApproachRecommendation],
                 iteration_history: List[Dict]):
        self.budget = budget
        self.recommendations = recommendations
        self.iteration_history = iteration_history

    def generate_text_report(self) -> str:
        """Generate a text-based calibration report."""
        lines = []

        # Header
        lines.append("=" * 80)
        lines.append("               FORECAST CALIBRATION REPORT")
        lines.append(f"               Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}")
        lines.append("=" * 80)
        lines.append("")

        # Iteration Summary
        if self.iteration_history:
            lines.append("OPTIMIZATION PROGRESS")
            lines.append("-" * 80)
            lines.append(f"{'Iteration':<12} {'Total Variance (£M)':<25}")
            lines.append("-" * 40)

            for hist in self.iteration_history:
                total_var = hist['total_variance'] / 1_000_000
                lines.append(f"{hist['iteration']:<12} {total_var:>20.2f}")

            lines.append("")

            # Show improvement
            if len(self.iteration_history) > 1:
                initial = self.iteration_history[0]['total_variance']
                final = self.iteration_history[-1]['total_variance']
                improvement = (initial - final) / 1_000_000
                lines.append(f"Total variance improvement: £{improvement:.2f}M")
                lines.append("")

        # Final Variance Summary
        if self.iteration_history:
            lines.append("FINAL VARIANCE SUMMARY")
            lines.append("-" * 80)

            final_summary = self.iteration_history[-1]['summary']
            lines.append(f"{'Metric':<20} {'Variance (£M)':<20} {'Status':<20}")
            lines.append("-" * 60)

            for metric, data in final_summary.items():
                var_m = data['total_variance'] / 1_000_000
                status = "OK" if abs(var_m) < Config.VARIANCE_TOLERANCE_M * 10 else "NEEDS ATTENTION"
                lines.append(f"{metric:<20} {var_m:>15.2f}     {status}")

            lines.append("")

        # Recommendations
        lines.append("RECOMMENDED CHANGES")
        lines.append("-" * 80)

        if not self.recommendations:
            lines.append("No changes recommended.")
        else:
            for i, rec in enumerate(self.recommendations, 1):
                lines.append(f"\nCHANGE {i}: {rec.segment} | {rec.metric} | {rec.cohort}")
                lines.append(f"  Current:     {rec.current_approach}" +
                           (f"({rec.current_param})" if rec.current_param else ""))
                lines.append(f"  Recommended: {rec.recommended_approach}" +
                           (f"({rec.recommended_param})" if rec.recommended_param else ""))
                lines.append(f"  Confidence:  {rec.confidence:.1%}")
                lines.append(f"  Rationale:   {rec.rationale}")

        lines.append("")
        lines.append("=" * 80)
        lines.append("                      END OF REPORT")
        lines.append("=" * 80)

        return "\n".join(lines)

    def save_report(self, filepath: str):
        """Save report to file."""
        report = self.generate_text_report()
        with open(filepath, 'w') as f:
            f.write(report)
        logger.info(f"Report saved to {filepath}")

    def generate_methodology_diff(self, original: pd.DataFrame, updated: pd.DataFrame) -> str:
        """Generate a diff showing methodology changes."""
        lines = []
        lines.append("METHODOLOGY CHANGES")
        lines.append("-" * 80)

        # Find differences
        for idx in range(len(updated)):
            row = updated.iloc[idx]

            # Find matching row in original
            mask = (
                (original['Segment'] == row['Segment']) &
                (original['Cohort'] == row['Cohort']) &
                (original['Metric'] == row['Metric']) &
                (original['MOB_Start'] == row['MOB_Start']) &
                (original['MOB_End'] == row['MOB_End'])
            )

            if not mask.any():
                # New rule
                lines.append(f"+ NEW: {row['Segment']}/{row['Cohort']}/{row['Metric']}: "
                           f"{row['Approach']}({row.get('Param1', '')})")
            else:
                orig_row = original[mask].iloc[0]
                if orig_row['Approach'] != row['Approach'] or str(orig_row.get('Param1', '')) != str(row.get('Param1', '')):
                    lines.append(f"~ CHANGED: {row['Segment']}/{row['Cohort']}/{row['Metric']}")
                    lines.append(f"    FROM: {orig_row['Approach']}({orig_row.get('Param1', '')})")
                    lines.append(f"    TO:   {row['Approach']}({row.get('Param1', '')})")

        return "\n".join(lines)


# =============================================================================
# SECTION 10: MAIN CALIBRATOR CLASS
# =============================================================================

class ForecastCalibrator:
    """Main calibrator class orchestrating the entire process."""

    def __init__(self, budget_file: str, methodology_file: str,
                 forecast_script: str, fact_raw_file: str):
        self.budget_file = budget_file
        self.methodology_file = methodology_file
        self.forecast_script = forecast_script
        self.fact_raw_file = fact_raw_file

        # Components
        self.budget_loader = BudgetLoader(budget_file)
        self.methodology_manager = MethodologyManager(methodology_file)
        self.forecast_runner = ForecastRunner(forecast_script, fact_raw_file)

        # Data
        self.budget = None
        self.recommendations = []
        self.iteration_history = []

    def run(self, mode: str = 'quick', max_iterations: int = None) -> Dict[str, Any]:
        """
        Run the calibration process.

        Args:
            mode: 'quick' for fast analysis, 'exhaustive' for full search
            max_iterations: Maximum optimization iterations

        Returns:
            Dict with results including recommendations, final variances, report
        """
        logger.info(f"Starting calibration in {mode} mode")

        # Load budget
        self.budget = self.budget_loader.load()

        # Load methodology
        self.methodology_manager.load()

        # Set iteration limit based on mode
        if max_iterations is None:
            max_iterations = 10 if mode == 'quick' else Config.MAX_ITERATIONS

        # Run optimization
        optimizer = OptimizationEngine(
            self.budget,
            self.methodology_manager,
            self.forecast_runner
        )

        self.recommendations = optimizer.optimize(max_iterations)
        self.iteration_history = optimizer.iteration_history

        # Generate report
        report_gen = ReportGenerator(
            self.budget,
            self.recommendations,
            self.iteration_history
        )

        report_text = report_gen.generate_text_report()

        return {
            'recommendations': self.recommendations,
            'iteration_history': self.iteration_history,
            'report': report_text,
            'final_methodology': self.methodology_manager.methodology
        }

    def compare_to_budget(self, forecast_file: str) -> Dict[str, Any]:
        """
        Compare a forecast output to budget without optimization.

        Useful for analyzing current state before running calibration.
        """
        # Load budget if not already loaded
        if self.budget is None:
            self.budget = self.budget_loader.load()

        # Load forecast
        loader = ForecastLoader()
        forecast = loader.load_from_combined_view(forecast_file)

        if forecast is None:
            return {'error': 'Failed to load forecast'}

        # Calculate variances
        calc = VarianceCalculator(self.budget, forecast)
        variances = calc.calculate_all_variances()
        summary = calc.get_summary_by_metric(variances)
        worst = calc.get_worst_variances(variances, n=20)

        return {
            'variances': variances,
            'summary': summary,
            'worst_variances': worst
        }

    def save_calibrated_methodology(self, filepath: str):
        """Save the calibrated methodology to a file."""
        if self.methodology_manager.methodology is not None:
            self.methodology_manager.save(filepath)
            logger.info(f"Calibrated methodology saved to {filepath}")


# =============================================================================
# SECTION 11: COMMAND LINE INTERFACE
# =============================================================================

def main():
    parser = argparse.ArgumentParser(
        description='Forecast Calibrator - Align forecast outputs with budget targets'
    )

    parser.add_argument(
        '--mode',
        choices=['quick', 'exhaustive', 'compare'],
        default='quick',
        help='Calibration mode: quick (fast), exhaustive (thorough), or compare (analysis only)'
    )

    parser.add_argument(
        '--budget',
        default=Config.BUDGET_FILE,
        help=f'Path to budget consol file (default: {Config.BUDGET_FILE})'
    )

    parser.add_argument(
        '--methodology',
        default=Config.METHODOLOGY_FILE,
        help=f'Path to rate methodology CSV (default: {Config.METHODOLOGY_FILE})'
    )

    parser.add_argument(
        '--forecast-script',
        default=Config.FORECAST_SCRIPT,
        help=f'Path to forecast script (default: {Config.FORECAST_SCRIPT})'
    )

    parser.add_argument(
        '--fact-raw',
        default=Config.FACT_RAW_FILE,
        help=f'Path to fact raw data (default: {Config.FACT_RAW_FILE})'
    )

    parser.add_argument(
        '--forecast-file',
        help='Existing forecast file for comparison mode'
    )

    parser.add_argument(
        '--max-iterations',
        type=int,
        help='Maximum optimization iterations'
    )

    parser.add_argument(
        '--output',
        default='calibration_output',
        help='Output directory for results'
    )

    parser.add_argument(
        '--verbose', '-v',
        action='store_true',
        help='Enable verbose logging'
    )

    args = parser.parse_args()

    # Set logging level
    if args.verbose:
        logging.getLogger().setLevel(logging.DEBUG)

    # Create output directory
    os.makedirs(args.output, exist_ok=True)

    # Initialize calibrator
    calibrator = ForecastCalibrator(
        budget_file=args.budget,
        methodology_file=args.methodology,
        forecast_script=args.forecast_script,
        fact_raw_file=args.fact_raw
    )

    if args.mode == 'compare':
        # Comparison mode - just analyze without optimization
        if not args.forecast_file:
            print("Error: --forecast-file required for compare mode")
            sys.exit(1)

        results = calibrator.compare_to_budget(args.forecast_file)

        if 'error' in results:
            print(f"Error: {results['error']}")
            sys.exit(1)

        # Print summary
        print("\n" + "=" * 100)
        print("                         FORECAST vs BUDGET COMPARISON REPORT")
        print(f"                         Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}")
        print("=" * 100)

        print("\n" + "=" * 80)
        print("VARIANCE SUMMARY BY METRIC (TOTAL ACROSS ALL MONTHS)")
        print("=" * 80)
        print(f"{'Metric':<20} {'Budget (£M)':<15} {'Forecast (£M)':<15} {'Variance (£M)':<15} {'Var %':<10}")
        print("-" * 75)

        for metric, data in results['summary'].items():
            budget_m = data['total_budget'] / 1_000_000
            forecast_m = data['total_forecast'] / 1_000_000
            variance_m = data['total_variance'] / 1_000_000
            var_pct = (variance_m / abs(budget_m) * 100) if abs(budget_m) > 0.001 else 0
            status = "OK" if abs(variance_m) < Config.VARIANCE_TOLERANCE_M * 10 else "NEEDS WORK"
            print(f"{metric:<20} {budget_m:>12.2f}   {forecast_m:>12.2f}   {variance_m:>12.2f}   {var_pct:>7.1f}%  {status}")

        # Detailed segment analysis
        print("\n" + "=" * 80)
        print("DETAILED SEGMENT-LEVEL ANALYSIS")
        print("=" * 80)

        # Get segment-level variances
        by_segment = calibrator.variance_calculator.get_variances_by_segment(results['variances']) if hasattr(calibrator, 'variance_calculator') else {}

        # Aggregate by segment and metric
        segment_summary = {}
        for v in results['variances']:
            key = (v.segment, v.metric)
            if key not in segment_summary:
                segment_summary[key] = {'budget': 0, 'forecast': 0, 'variance': 0}
            segment_summary[key]['budget'] += v.budget_value
            segment_summary[key]['forecast'] += v.forecast_value
            segment_summary[key]['variance'] += v.variance

        # Print by metric
        for metric in ['Collections', 'ClosingGBV', 'ClosingNBV', 'Revenue', 'NetImpairment']:
            print(f"\n{metric}:")
            print(f"  {'Segment':<25} {'Budget (£M)':<12} {'Forecast (£M)':<12} {'Variance (£M)':<12} {'Var %':<8}")
            print("  " + "-" * 70)

            for segment in ['Non Prime', 'Near Prime Small', 'Near Prime Medium', 'Prime']:
                key = (segment, metric)
                if key in segment_summary:
                    data = segment_summary[key]
                    budget_m = data['budget'] / 1_000_000
                    forecast_m = data['forecast'] / 1_000_000
                    variance_m = data['variance'] / 1_000_000
                    var_pct = (variance_m / abs(budget_m) * 100) if abs(budget_m) > 0.001 else 0
                    print(f"  {segment:<25} {budget_m:>10.2f}   {forecast_m:>10.2f}   {variance_m:>10.2f}   {var_pct:>6.1f}%")

        print("\n" + "=" * 80)
        print("TOP 10 WORST MONTHLY VARIANCES")
        print("=" * 80)
        print(f"{'Metric':<15} {'Segment':<22} {'Month':<10} {'Variance (£M)':<15}")
        print("-" * 65)
        for v in results['worst_variances'][:10]:
            var_m = v.variance / 1_000_000
            print(f"{v.metric:<15} {v.segment:<22} {v.month:<10} {var_m:>12.2f}")

        # Save detailed report to file
        report_path = os.path.join(args.output, 'variance_comparison_report.txt')
        os.makedirs(args.output, exist_ok=True)

        with open(report_path, 'w') as f:
            f.write("FORECAST vs BUDGET VARIANCE COMPARISON REPORT\n")
            f.write(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}\n")
            f.write("=" * 80 + "\n\n")

            f.write("SUMMARY BY METRIC:\n")
            for metric, data in results['summary'].items():
                budget_m = data['total_budget'] / 1_000_000
                forecast_m = data['total_forecast'] / 1_000_000
                variance_m = data['total_variance'] / 1_000_000
                var_pct = (variance_m / abs(budget_m) * 100) if abs(budget_m) > 0.001 else 0
                f.write(f"  {metric}: Budget=£{budget_m:.2f}M, Forecast=£{forecast_m:.2f}M, Variance=£{variance_m:.2f}M ({var_pct:.1f}%)\n")

            f.write("\n\nDETAILED BY SEGMENT AND MONTH:\n")
            for v in sorted(results['variances'], key=lambda x: (x.metric, x.segment, x.month)):
                var_m = v.variance / 1_000_000
                f.write(f"  {v.metric}/{v.segment}/{v.month}: Budget=£{v.budget_value/1e6:.2f}M, Forecast=£{v.forecast_value/1e6:.2f}M, Variance=£{var_m:.2f}M\n")

        print(f"\n\nDetailed report saved to: {report_path}")

    else:
        # Optimization mode
        results = calibrator.run(
            mode=args.mode,
            max_iterations=args.max_iterations
        )

        # Print report
        print(results['report'])

        # Save outputs
        report_path = os.path.join(args.output, 'calibration_report.txt')
        with open(report_path, 'w') as f:
            f.write(results['report'])
        print(f"\nReport saved to: {report_path}")

        # Save calibrated methodology
        meth_path = os.path.join(args.output, 'Rate_Methodology_Calibrated.csv')
        calibrator.save_calibrated_methodology(meth_path)
        print(f"Calibrated methodology saved to: {meth_path}")


if __name__ == '__main__':
    main()
