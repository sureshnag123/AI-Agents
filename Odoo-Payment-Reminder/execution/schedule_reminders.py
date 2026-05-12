#!/usr/bin/env python3
"""
Payment Reminder Scheduler

Runs payment reminders on a configurable schedule using APScheduler.
Can also be used as a one-shot cron replacement via --once flag.

Usage:
    # Run scheduler daemon (sends reminders daily at configured time)
    python schedule_reminders.py

    # Run once and exit (for Windows Task Scheduler / cron)
    python schedule_reminders.py --once

    # Custom schedule
    python schedule_reminders.py --hour 9 --minute 0

    # Run with escalation (different emails at 7, 14, 30 days overdue)
    python schedule_reminders.py --once --escalate
"""

import os
import sys
import argparse
import logging
from datetime import date

from dotenv import load_dotenv

load_dotenv()

# Add execution directory to path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler(
            os.path.join(
                os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                ".tmp",
                "scheduler.log",
            ),
            mode="a",
        ),
    ],
)

logger = logging.getLogger(__name__)


def run_reminder_cycle(escalate: bool = False, use_smtp: bool = False):
    """Execute one full reminder cycle."""
    from send_payment_reminders import process_reminders

    logger.info("=" * 60)
    logger.info(f"Starting reminder cycle — {date.today()}")
    logger.info("=" * 60)

    try:
        if escalate:
            # Escalation: send different reminders based on how overdue
            # Level 1: Due in 3 days (gentle pre-reminder)
            logger.info("Phase 1: Pre-due reminders (next 3 days)...")
            process_reminders(mode="upcoming", days=3, use_smtp=use_smtp)

            # Level 2: 1-7 days overdue (friendly reminder)
            logger.info("Phase 2: Recently overdue (1-7 days)...")
            process_reminders(mode="overdue", days=0, use_smtp=use_smtp)

            # Level 3: 14+ days overdue (firm reminder)
            logger.info("Phase 3: Significantly overdue (14+ days)...")
            process_reminders(mode="overdue", days=14, use_smtp=use_smtp)

            # Level 4: 30+ days overdue (urgent)
            logger.info("Phase 4: Urgently overdue (30+ days)...")
            process_reminders(mode="overdue", days=30, use_smtp=use_smtp)
        else:
            # Simple mode: upcoming + all overdue
            logger.info("Sending upcoming-due reminders (next 7 days)...")
            process_reminders(mode="upcoming", days=7, use_smtp=use_smtp)

            logger.info("Sending overdue reminders...")
            process_reminders(mode="overdue", days=0, use_smtp=use_smtp)

        logger.info("Reminder cycle complete.")

    except Exception as e:
        logger.error(f"Reminder cycle failed: {e}", exc_info=True)


def start_scheduler(hour: int = 9, minute: int = 0, escalate: bool = False, use_smtp: bool = False):
    """Start APScheduler to run reminders daily."""
    try:
        from apscheduler.schedulers.blocking import BlockingScheduler
        from apscheduler.triggers.cron import CronTrigger
    except ImportError:
        logger.error(
            "APScheduler not installed. Install with: pip install apscheduler\n"
            "Or use --once flag with Windows Task Scheduler / cron instead."
        )
        sys.exit(1)

    scheduler = BlockingScheduler()

    scheduler.add_job(
        run_reminder_cycle,
        CronTrigger(hour=hour, minute=minute),
        kwargs={"escalate": escalate, "use_smtp": use_smtp},
        id="payment_reminders",
        name="Daily Payment Reminders",
        replace_existing=True,
    )

    logger.info(f"Scheduler started — reminders will run daily at {hour:02d}:{minute:02d}")
    logger.info("Press Ctrl+C to stop")

    try:
        scheduler.start()
    except (KeyboardInterrupt, SystemExit):
        logger.info("Scheduler stopped.")


def main():
    parser = argparse.ArgumentParser(
        description="Schedule and run Odoo payment reminders"
    )
    parser.add_argument(
        "--once",
        action="store_true",
        help="Run once and exit (for Task Scheduler / cron)",
    )
    parser.add_argument(
        "--hour",
        type=int,
        default=int(os.getenv("REMINDER_HOUR", "9")),
        help="Hour to send reminders (0-23, default: 9)",
    )
    parser.add_argument(
        "--minute",
        type=int,
        default=int(os.getenv("REMINDER_MINUTE", "0")),
        help="Minute to send reminders (0-59, default: 0)",
    )
    parser.add_argument(
        "--escalate",
        action="store_true",
        help="Use escalation mode (different emails by overdue severity)",
    )
    parser.add_argument(
        "--smtp",
        action="store_true",
        help="Send via SMTP instead of Odoo mail",
    )
    args = parser.parse_args()

    # Ensure .tmp directory exists for logs
    tmp_dir = os.path.join(
        os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
        ".tmp",
    )
    os.makedirs(tmp_dir, exist_ok=True)

    if args.once:
        run_reminder_cycle(escalate=args.escalate, use_smtp=args.smtp)
    else:
        start_scheduler(
            hour=args.hour,
            minute=args.minute,
            escalate=args.escalate,
            use_smtp=args.smtp,
        )


if __name__ == "__main__":
    main()
