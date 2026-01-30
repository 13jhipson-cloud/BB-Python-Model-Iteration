#!/usr/bin/env python3
"""
Impairment Calibrator - Iteratively adjusts coverage ratios to match budget impairment.

This script:
1. Reads budget impairment targets
2. Runs forecasts with different coverage ratio configurations
3. Compares forecast impairment to budget
4. Adjusts coverage ratios until variance is minimized
"""

import pandas as pd
import numpy as np
import subprocess
import os
import sys
import logging
from datetime import datetime
from typing import Dict, List, Tuple, Optional
import shutil
import tempfile

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)


class BudgetLoader:
    """Load and parse budget data."""

    def __init__(self, budget_file: str):
        self.budget_file = budget_file
        self.budget_data = None

    def load(self) -> pd.DataFrame:
        """Load budget impairment data."""
        df = pd.read_excel(self.budget_file, 'P&L analysis - BB', header=None)

        # Get dates from row 2 (columns 3 onwards)
        dates = df.iloc[2, 4:].tolist()  # Skip first few columns

        # Get gross impairment from row 77
        gross_imp = df.iloc[77, 4:].tolist()

        # Build budget dataframe
        budget_records = []
        for d, imp in zip(dates, gross_imp):
            if pd.notna(d) and pd.notna(imp):
                try:
                    if isinstance(d, str):
                        dt = pd.to_datetime(d)
                    else:
                        dt = pd.to_datetime(d)
                    budget_records.append({
                        'Month': dt,
                        'Budget_Impairment': float(imp)
                    })
                except:
                    continue

        self.budget_data = pd.DataFrame(budget_records)
        logger.info(f"Loaded {len(self.budget_data)} months of budget data")
        return self.budget_data


class ForecastRunner:
    """Run the forecast model."""

    def __init__(self, forecast_script: str = 'backbook_forecast.py', fact_raw_file: str = 'Fact_Raw_New.xlsx'):
        self.forecast_script = forecast_script
        self.fact_raw_file = fact_raw_file

    def run(self, methodology_file: str, output_dir: str) -> Optional[str]:
        """Run forecast and return path to transparency report."""
        os.makedirs(output_dir, exist_ok=True)

        cmd = [
            sys.executable, self.forecast_script,
            '--fact-raw', self.fact_raw_file,
            '--methodology', methodology_file,
            '--output', output_dir,
            '--transparency-report'
        ]

        logger.info(f"Running forecast...")
        try:
            result = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                timeout=300
            )

            if result.returncode != 0:
                logger.error(f"Forecast failed: {result.stderr[:500]}")
                return None

            report_path = os.path.join(output_dir, 'Forecast_Transparency_Report.xlsx')
            if os.path.exists(report_path):
                return report_path

            logger.error(f"Report not found at {report_path}")
            return None

        except subprocess.TimeoutExpired:
            logger.error("Forecast timed out")
            return None
        except Exception as e:
            logger.error(f"Error running forecast: {e}")
            return None


class ImpairmentAnalyzer:
    """Analyze forecast impairment vs budget."""

    def __init__(self):
        pass

    def load_forecast(self, report_path: str) -> pd.DataFrame:
        """Load forecast impairment data from transparency report."""
        df = pd.read_excel(report_path, '9_Summary')

        # Aggregate by month
        df['ForecastMonth'] = pd.to_datetime(df['ForecastMonth'])
        monthly = df.groupby('ForecastMonth').agg({
            'Net_Impairment': 'sum',
            'ClosingGBV': 'sum',
            'Total_Coverage_Ratio': 'mean'
        }).reset_index()

        monthly.columns = ['Month', 'Forecast_Impairment', 'ClosingGBV', 'Avg_Coverage_Ratio']
        return monthly

    def compare(self, forecast_df: pd.DataFrame, budget_df: pd.DataFrame) -> pd.DataFrame:
        """Compare forecast to budget."""
        # Merge on month
        merged = pd.merge(
            forecast_df,
            budget_df,
            on='Month',
            how='inner'
        )

        merged['Variance'] = merged['Forecast_Impairment'] - merged['Budget_Impairment']
        merged['Variance_Pct'] = merged['Variance'] / merged['Budget_Impairment'].abs() * 100

        return merged


class MethodologyGenerator:
    """Generate rate methodology with specific coverage ratios."""

    # Base methodology template (non-coverage ratio entries)
    BASE_METHODOLOGY = """Segment,Cohort,Metric,MOB_Start,MOB_End,Approach,Param1,Param2,Explanation
NRP-S,ALL,Coll_Principal,0,12,DonorCohort,202405,,Young-donor
NRP-S,ALL,Coll_Principal,13,36,CohortAvg,6,,Mid-6mo avg
NRP-S,ALL,Coll_Principal,37,999,CohortAvg,6,,Mature-6mo
NRP-L,ALL,Coll_Interest,0,12,DonorCohort,202405,,Young-donor
NRP-L,ALL,Coll_Interest,13,999,CohortAvg,6,,Mid/mature-6mo
NRP-M,ALL,InterestRevenue,0,12,DonorCohort,202405,,Young-donor
NRP-M,ALL,InterestRevenue,13,36,CohortTrend,6,,Mid-trend
NRP-M,ALL,InterestRevenue,37,999,CohortAvg,6,,Mature-flat
NRP-M,ALL,WO_Other,0,999,CohortAvg,12,,Rare-12mo
NRP-M,ALL,WO_DebtSold,0,999,SegMedian,,,Portfolio decision
NRP-M,ALL,Debt_Sale_Coverage_Ratio,0,999,Manual,0.785,,DS coverage
NRP-M,ALL,Debt_Sale_Proceeds_Rate,0,999,Manual,0.24,,Proceeds 24p
NON PRIME,ALL,Coll_Principal,0,12,DonorCohort,202404,,Young-donor
NON PRIME,ALL,Coll_Principal,13,36,CohortAvg,6,,Mid-6mo avg
NON PRIME,ALL,Coll_Principal,37,999,CohortAvg,6,,Mature-6mo
NON PRIME,ALL,Coll_Interest,0,12,DonorCohort,202404,,Young-donor
NON PRIME,ALL,Coll_Interest,13,999,CohortAvg,6,,Mid/mature-6mo
NON PRIME,ALL,InterestRevenue,0,12,DonorCohort,202404,,Young-donor
NON PRIME,ALL,InterestRevenue,13,36,CohortTrend,6,,Mid-trend
NON PRIME,ALL,InterestRevenue,37,999,CohortAvg,6,,Mature-flat
NON PRIME,ALL,WO_Other,0,999,CohortAvg,12,,Rare-12mo
NON PRIME,ALL,WO_DebtSold,0,999,SegMedian,,,Portfolio decision
NON PRIME,ALL,Debt_Sale_Coverage_Ratio,0,999,Manual,0.785,,DS coverage
NON PRIME,ALL,Debt_Sale_Proceeds_Rate,0,999,Manual,0.24,,Proceeds 24p
PRIME,ALL,Coll_Principal,0,12,DonorCohort,202405,,Young-donor
PRIME,ALL,Coll_Principal,13,36,CohortAvg,6,,Mid-6mo avg
PRIME,ALL,Coll_Principal,37,999,CohortAvg,6,,Mature-6mo
PRIME,ALL,Coll_Interest,0,12,DonorCohort,202405,,Young-donor
PRIME,ALL,Coll_Interest,13,999,CohortAvg,6,,Mid/mature-6mo
PRIME,ALL,InterestRevenue,0,12,DonorCohort,202405,,Young-donor
PRIME,ALL,InterestRevenue,13,36,CohortTrend,6,,Mid-trend
PRIME,ALL,InterestRevenue,37,999,CohortAvg,6,,Mature-flat
PRIME,ALL,WO_Other,0,999,CohortAvg,12,,Rare-12mo
PRIME,ALL,WO_DebtSold,0,999,SegMedian,,,Portfolio decision
PRIME,ALL,Debt_Sale_Coverage_Ratio,0,999,Manual,0.785,,DS coverage
PRIME,ALL,Debt_Sale_Proceeds_Rate,0,999,Manual,0.24,,Proceeds 24p
NRP-S,ALL,InterestRevenue,0,12,DonorCohort,202405,,Young-donor
NRP-S,ALL,InterestRevenue,13,36,CohortTrend,6,,Mid-trend
NRP-S,ALL,InterestRevenue,37,999,CohortAvg,6,,Mature-flat
NRP-L,ALL,WO_Other,0,999,CohortAvg,12,,Rare-12mo
NRP-L,ALL,WO_DebtSold,0,999,SegMedian,,,Portfolio decision
NRP-L,ALL,Debt_Sale_Coverage_Ratio,0,999,Manual,0.785,,DS coverage
NRP-L,ALL,Debt_Sale_Proceeds_Rate,0,999,Manual,0.24,,Proceeds 24p
NRP-S,ALL,WO_Other,0,999,CohortAvg,12,,Rare-12mo
NRP-S,ALL,WO_DebtSold,0,999,SegMedian,,,Portfolio decision
NRP-S,ALL,Debt_Sale_Coverage_Ratio,0,999,Manual,0.785,,DS coverage
NRP-S,ALL,Debt_Sale_Proceeds_Rate,0,999,Manual,0.24,,Proceeds 24p
NRP-M,ALL,Coll_Principal,0,12,DonorCohort,202405,,Young-donor
NRP-M,ALL,Coll_Principal,13,36,CohortAvg,6,,Mid-6mo avg
NRP-M,ALL,Coll_Principal,37,999,CohortAvg,6,,Mature-6mo
NRP-M,ALL,Coll_Interest,0,12,DonorCohort,202405,,Young-donor
NRP-M,ALL,Coll_Interest,13,999,CohortAvg,6,,Mid/mature-6mo
NRP-L,ALL,Coll_Interest,0,12,DonorCohort,202405,,Young-donor
NRP-L,ALL,Coll_Interest,13,999,CohortAvg,6,,Mid/mature-6mo
NRP-L,ALL,Coll_Principal,0,12,DonorCohort,202405,,Young-donor
NRP-L,ALL,Coll_Principal,13,36,CohortAvg,6,,Mid-6mo avg
NRP-L,ALL,Coll_Principal,37,999,CohortAvg,6,,Mature-6mo
NRP-L,ALL,InterestRevenue,0,12,DonorCohort,202405,,Young-donor
NRP-L,ALL,InterestRevenue,13,36,CohortTrend,6,,Mid-trend
NRP-L,ALL,InterestRevenue,37,999,CohortAvg,6,,Mature-flat"""

    SEGMENTS = ['NON PRIME', 'NRP-S', 'NRP-M', 'NRP-L', 'PRIME']

    def __init__(self):
        pass

    def generate(self, coverage_ratios: Dict[str, float], output_file: str) -> str:
        """
        Generate methodology file with specified coverage ratios.

        Args:
            coverage_ratios: Dict mapping segment -> coverage ratio (e.g., {'NON PRIME': 0.20})
            output_file: Path to output CSV file

        Returns:
            Path to generated file
        """
        lines = self.BASE_METHODOLOGY.strip().split('\n')

        # Add coverage ratio entries for each segment
        for segment in self.SEGMENTS:
            ratio = coverage_ratios.get(segment, 0.15)  # Default 15%
            lines.append(f"{segment},ALL,Total_Coverage_Ratio,0,999,Manual,{ratio:.4f},,Calibrated coverage ratio")

        with open(output_file, 'w') as f:
            f.write('\n'.join(lines))

        return output_file


class ImpairmentCalibrator:
    """Main calibration orchestrator."""

    def __init__(self, budget_file: str = 'Budget consol file.xlsx'):
        self.budget_loader = BudgetLoader(budget_file)
        self.forecast_runner = ForecastRunner()
        self.analyzer = ImpairmentAnalyzer()
        self.methodology_gen = MethodologyGenerator()

        # Initial coverage ratio guesses by segment
        # Starting with lower values since high ratios cause too much release
        self.coverage_ratios = {
            'NON PRIME': 0.08,   # Higher risk segment
            'NRP-S': 0.05,       # Near prime small
            'NRP-M': 0.04,       # Near prime medium
            'NRP-L': 0.05,       # Near prime large
            'PRIME': 0.01        # Low risk segment
        }

        self.iteration_history = []

    def run_iteration(self, iteration: int, output_base: str) -> Dict:
        """Run a single calibration iteration."""
        logger.info(f"\n{'='*60}")
        logger.info(f"ITERATION {iteration}")
        logger.info(f"Coverage ratios: {self.coverage_ratios}")
        logger.info(f"{'='*60}")

        # Generate methodology
        methodology_file = os.path.join(output_base, f'methodology_iter{iteration}.csv')
        self.methodology_gen.generate(self.coverage_ratios, methodology_file)

        # Run forecast
        output_dir = os.path.join(output_base, f'forecast_iter{iteration}')
        report_path = self.forecast_runner.run(methodology_file, output_dir)

        if not report_path:
            logger.error("Forecast failed")
            return {'success': False, 'iteration': iteration}

        # Load and compare
        forecast_df = self.analyzer.load_forecast(report_path)
        budget_df = self.budget_loader.budget_data
        comparison = self.analyzer.compare(forecast_df, budget_df)

        # Calculate metrics
        total_variance = comparison['Variance'].sum()
        avg_variance = comparison['Variance'].mean()
        max_variance = comparison['Variance'].abs().max()

        result = {
            'success': True,
            'iteration': iteration,
            'coverage_ratios': self.coverage_ratios.copy(),
            'total_variance': total_variance,
            'avg_variance': avg_variance,
            'max_variance': max_variance,
            'comparison': comparison,
            'methodology_file': methodology_file,
            'report_path': report_path
        }

        self.iteration_history.append(result)

        logger.info(f"Total Variance: £{total_variance:,.0f}")
        logger.info(f"Average Monthly Variance: £{avg_variance:,.0f}")
        logger.info(f"Max Monthly Variance: £{max_variance:,.0f}")

        # Show first few months
        logger.info("\nMonthly Comparison (first 6 months):")
        for _, row in comparison.head(6).iterrows():
            logger.info(f"  {row['Month'].strftime('%Y-%m')}: "
                       f"Forecast £{row['Forecast_Impairment']:,.0f} | "
                       f"Budget £{row['Budget_Impairment']:,.0f} | "
                       f"Var £{row['Variance']:,.0f}")

        return result

    def adjust_coverage_ratios(self, result: Dict) -> bool:
        """
        Adjust coverage ratios based on variance analysis.

        Returns True if adjustments were made, False if converged.
        """
        comparison = result['comparison']

        # If average variance is small enough, we're done
        if abs(result['avg_variance']) < 200000:  # £200k tolerance
            logger.info("Converged! Average variance within tolerance.")
            return False

        # The impairment mechanics are:
        # - Budget expects NEGATIVE impairment (provision charge = increasing provisions)
        # - If forecast shows POSITIVE impairment, provisions are being released
        # - To make impairment MORE NEGATIVE, we need LOWER coverage ratios
        #   (so provisions start lower and don't release as much during runoff)
        #
        # So if variance > 0 (forecast too positive), DECREASE coverage ratios
        # If variance < 0 (forecast too negative), INCREASE coverage ratios

        avg_variance = result['avg_variance']
        avg_budget = comparison['Budget_Impairment'].mean()

        # Proportional adjustment - INVERTED from before
        # If variance > 0, we need factor < 1 to decrease ratios
        adjustment_factor = 1.0 - (avg_variance / abs(avg_budget)) * 0.2
        adjustment_factor = max(0.7, min(1.3, adjustment_factor))  # Clamp to ±30%

        logger.info(f"\nAdjustment factor: {adjustment_factor:.3f}")

        # Apply adjustment to all segments
        for segment in self.coverage_ratios:
            old_ratio = self.coverage_ratios[segment]
            new_ratio = old_ratio * adjustment_factor
            new_ratio = max(0.01, min(0.50, new_ratio))  # Clamp between 1% and 50%
            self.coverage_ratios[segment] = round(new_ratio, 4)
            logger.info(f"  {segment}: {old_ratio:.4f} -> {new_ratio:.4f}")

        return True

    def calibrate(self, max_iterations: int = 10, output_dir: str = 'calibration_output') -> Dict:
        """
        Run the full calibration process.

        Args:
            max_iterations: Maximum number of iterations
            output_dir: Directory to save outputs

        Returns:
            Final result dictionary
        """
        logger.info("Starting Impairment Calibration")
        logger.info(f"Max iterations: {max_iterations}")

        # Load budget
        self.budget_loader.load()

        # Create output directory
        os.makedirs(output_dir, exist_ok=True)

        best_result = None
        best_variance = float('inf')

        for iteration in range(1, max_iterations + 1):
            result = self.run_iteration(iteration, output_dir)

            if not result['success']:
                logger.error(f"Iteration {iteration} failed, stopping.")
                break

            # Track best result
            if abs(result['avg_variance']) < best_variance:
                best_variance = abs(result['avg_variance'])
                best_result = result

            # Check convergence and adjust
            if not self.adjust_coverage_ratios(result):
                logger.info("Calibration converged!")
                break

        # Save final methodology
        if best_result:
            final_methodology = os.path.join(output_dir, 'Rate_Methodology_Calibrated.csv')
            shutil.copy(best_result['methodology_file'], final_methodology)
            logger.info(f"\nFinal calibrated methodology saved to: {final_methodology}")

            # Generate summary report
            self.generate_report(output_dir, best_result)

        return best_result

    def generate_report(self, output_dir: str, best_result: Dict):
        """Generate calibration summary report."""
        report_lines = [
            "IMPAIRMENT CALIBRATION REPORT",
            f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}",
            "=" * 60,
            "",
            "FINAL COVERAGE RATIOS:",
        ]

        for segment, ratio in best_result['coverage_ratios'].items():
            report_lines.append(f"  {segment}: {ratio:.2%}")

        report_lines.extend([
            "",
            "VARIANCE SUMMARY:",
            f"  Total Variance: £{best_result['total_variance']:,.0f}",
            f"  Average Monthly: £{best_result['avg_variance']:,.0f}",
            f"  Max Monthly: £{best_result['max_variance']:,.0f}",
            "",
            "MONTHLY COMPARISON:",
        ])

        for _, row in best_result['comparison'].iterrows():
            report_lines.append(
                f"  {row['Month'].strftime('%Y-%m')}: "
                f"Forecast £{row['Forecast_Impairment']:,.0f} | "
                f"Budget £{row['Budget_Impairment']:,.0f} | "
                f"Variance £{row['Variance']:,.0f}"
            )

        report_path = os.path.join(output_dir, 'calibration_report.txt')
        with open(report_path, 'w') as f:
            f.write('\n'.join(report_lines))

        logger.info(f"Report saved to: {report_path}")


def main():
    import argparse

    parser = argparse.ArgumentParser(description='Calibrate coverage ratios to match budget impairment')
    parser.add_argument('--max-iterations', type=int, default=10, help='Maximum iterations')
    parser.add_argument('--output', type=str, default='impairment_calibration', help='Output directory')
    parser.add_argument('--budget-file', type=str, default='Budget consol file.xlsx', help='Budget file path')

    args = parser.parse_args()

    calibrator = ImpairmentCalibrator(budget_file=args.budget_file)
    result = calibrator.calibrate(max_iterations=args.max_iterations, output_dir=args.output)

    if result:
        print(f"\nCalibration complete!")
        print(f"Final methodology: {args.output}/Rate_Methodology_Calibrated.csv")
        print(f"Average variance: £{result['avg_variance']:,.0f}")
    else:
        print("Calibration failed.")


if __name__ == '__main__':
    main()
