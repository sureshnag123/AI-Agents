"""
CLI Runner for GSTR2B Reconciliation.

Usage:
    python run.py --tally path/to/tally.xlsx --gstr2b path/to/gstr2b.xlsx --month 2026-03
"""
import argparse
import logging
import sys
from pathlib import Path
from datetime import datetime

import config
from core.data_loader import load_tally_data, load_gstr2b_data
from core.reconciler import Reconciler
from core.report_generator import generate_reports


def setup_logging(log_dir: Path):
    """Configure logging to file and console."""
    log_dir.mkdir(parents=True, exist_ok=True)
    log_file = log_dir / f"reco_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"

    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s | %(levelname)-8s | %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
        handlers=[
            logging.FileHandler(log_file, encoding="utf-8"),
            logging.StreamHandler(sys.stdout),
        ],
    )
    return log_file


def main():
    parser = argparse.ArgumentParser(description="GSTR2B Reconciliation Agent")
    parser.add_argument("--tally", required=True, help="Path to Tally export file (xlsx/csv)")
    parser.add_argument("--gstr2b", required=True, help="Path to GSTR2B file (xlsx/csv)")
    parser.add_argument("--month", default="", help="Month in YYYY-MM format (default: current)")
    parser.add_argument("--output", default="", help="Output directory (default: auto)")
    args = parser.parse_args()

    # Setup
    month_str = args.month or datetime.now().strftime("%Y-%m")
    output_dir = Path(args.output) if args.output else config.OUTPUT_DIR / month_str
    log_file = setup_logging(config.LOG_DIR)

    logger = logging.getLogger("gstr2b_reco")
    logger.info("=" * 60)
    logger.info("GSTR2B RECONCILIATION AGENT")
    logger.info(f"  Tally file : {args.tally}")
    logger.info(f"  GSTR2B file: {args.gstr2b}")
    logger.info(f"  Month      : {month_str}")
    logger.info(f"  Output dir : {output_dir}")
    logger.info(f"  Log file   : {log_file}")
    logger.info("=" * 60)

    try:
        # Load data
        tally_df = load_tally_data(args.tally)
        gstr2b_df = load_gstr2b_data(args.gstr2b)

        # Run reconciliation
        reconciler = Reconciler(tally_df, gstr2b_df)
        result = reconciler.run()

        # Generate reports
        files = generate_reports(result, output_dir, month_str)

        print("\n" + "=" * 60)
        print("RECONCILIATION COMPLETE")
        print("=" * 60)
        print(f"  Matched           : {result.summary['matched_count']}")
        print(f"  In Tally only     : {result.summary['in_tally_not_gstr2b']}")
        print(f"  In GSTR2B only    : {result.summary['in_gstr2b_not_tally']}")
        print(f"  Amount mismatches : {result.summary['amount_mismatch']}")
        print(f"  Time taken        : {result.summary['run_time_seconds']}s")
        print(f"\nReports saved to: {output_dir}")
        for name, path in files.items():
            print(f"  {name}: {path}")

    except Exception as e:
        logger.exception(f"Reconciliation failed: {e}")
        print(f"\nERROR: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
