#!/usr/bin/env python3
"""
Sync Scheduler

Runs the Odoo → Tally sync on a daily or weekly schedule.

Usage:
    # Daily at 8 PM
    python sync_scheduler.py --hour 20 --minute 0

    # Weekly on Saturday at 10 AM
    python sync_scheduler.py --weekly --day saturday --hour 10

    # One-shot (for Windows Task Scheduler / cron)
    python sync_scheduler.py --once

    # One-shot with specific voucher types
    python sync_scheduler.py --once --types "purchase,sales"
"""

import os
import sys
import argparse
import logging
from datetime import datetime
from pathlib import Path

from dotenv import load_dotenv

SCRIPT_DIR = Path(__file__).parent.resolve()
PROJECT_ROOT = SCRIPT_DIR.parent
TMP_DIR = PROJECT_ROOT / ".tmp"

load_dotenv(PROJECT_ROOT / ".env")
sys.path.insert(0, str(SCRIPT_DIR))

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
)
log = logging.getLogger("sync_scheduler")


def run_sync(types: str = None, mode: str = "incremental"):
    """Execute a single sync run."""
    from odoo_connector import OdooConnector
    from tally_connector import TallyConnector
    from tally_xml_builder import TallyXMLBuilder
    from odoo_tally_sync import sync_vouchers, VOUCHER_TYPE_MAP

    log.info(f"Starting scheduled sync at {datetime.now().isoformat()}")

    try:
        odoo = OdooConnector()
        company_name = os.getenv("TALLY_COMPANY_NAME", "")
        tally = TallyConnector()
        builder = TallyXMLBuilder(company_name=company_name)

        # Test connections
        tally.test_connection()

        if types:
            voucher_types = [t.strip().lower() for t in types.split(",")]
        else:
            voucher_types = list(VOUCHER_TYPE_MAP.keys())

        result = sync_vouchers(
            odoo=odoo,
            tally=tally,
            builder=builder,
            voucher_types=voucher_types,
            mode=mode,
        )

        totals = result.get("totals", {})
        log.info(f"Sync complete. Synced: {totals.get('synced', 0)}, "
                 f"Errors: {totals.get('errors', 0)}")

    except Exception as e:
        log.error(f"Scheduled sync failed: {e}", exc_info=True)


def main():
    parser = argparse.ArgumentParser(description="Odoo-Tally Sync Scheduler")
    parser.add_argument("--once", action="store_true",
                        help="Run once and exit (for Task Scheduler / cron)")
    parser.add_argument("--weekly", action="store_true",
                        help="Run weekly instead of daily")
    parser.add_argument("--day", type=str, default="saturday",
                        help="Day of week for weekly schedule (default: saturday)")
    parser.add_argument("--hour", type=int,
                        default=int(os.getenv("SYNC_HOUR", "20")),
                        help="Hour to run (0-23, default: 20)")
    parser.add_argument("--minute", type=int,
                        default=int(os.getenv("SYNC_MINUTE", "0")),
                        help="Minute to run (0-59, default: 0)")
    parser.add_argument("--types", type=str, default=None,
                        help="Comma-separated voucher types to sync")
    parser.add_argument("--mode", choices=["incremental", "full"],
                        default="incremental",
                        help="Sync mode (default: incremental)")

    args = parser.parse_args()
    TMP_DIR.mkdir(parents=True, exist_ok=True)

    if args.once:
        log.info("One-shot sync mode")
        run_sync(types=args.types, mode=args.mode)
        return

    # APScheduler-based recurring schedule
    try:
        from apscheduler.schedulers.blocking import BlockingScheduler
        from apscheduler.triggers.cron import CronTrigger
    except ImportError:
        log.error("APScheduler required for recurring schedules. "
                  "Install with: pip install apscheduler")
        log.info("Alternatively, use --once with Windows Task Scheduler or cron.")
        sys.exit(1)

    scheduler = BlockingScheduler()

    if args.weekly:
        trigger = CronTrigger(
            day_of_week=args.day[:3].lower(),
            hour=args.hour,
            minute=args.minute,
        )
        schedule_desc = f"Weekly on {args.day} at {args.hour:02d}:{args.minute:02d}"
    else:
        trigger = CronTrigger(
            hour=args.hour,
            minute=args.minute,
        )
        schedule_desc = f"Daily at {args.hour:02d}:{args.minute:02d}"

    scheduler.add_job(
        run_sync,
        trigger=trigger,
        kwargs={"types": args.types, "mode": args.mode},
        id="odoo_tally_sync",
        name="Odoo → Tally Sync",
    )

    log.info(f"Scheduler started — {schedule_desc}")
    log.info("Press Ctrl+C to stop")

    try:
        scheduler.start()
    except KeyboardInterrupt:
        log.info("Scheduler stopped by user")
    except Exception as e:
        log.error(f"Scheduler error: {e}")


if __name__ == "__main__":
    main()
