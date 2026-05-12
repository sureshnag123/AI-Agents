#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Repair bare Tally ledgers — fetch full details from Odoo and ALTER in Tally.

Targets ledgers that were created with only a name (no GST, address, state,
pincode) due to a now-fixed bug in the XML field names.

Usage:
    python repair_ledger_details.py               # dry run — show what would be updated
    python repair_ledger_details.py --execute     # actually ALTER ledgers in Tally
"""

import io
import os
import re
import sys
import argparse
import logging
from pathlib import Path

if sys.stdout.encoding and sys.stdout.encoding.lower() != "utf-8":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
if sys.stderr.encoding and sys.stderr.encoding.lower() != "utf-8":
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")

from dotenv import load_dotenv

SCRIPT_DIR   = Path(__file__).parent.resolve()
PROJECT_ROOT = SCRIPT_DIR.parent
load_dotenv(PROJECT_ROOT / ".env")
sys.path.insert(0, str(SCRIPT_DIR))

from odoo_connector  import OdooConnector
from tally_connector import TallyConnector

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.StreamHandler()],
)
log = logging.getLogger("repair_ledgers")

# ── Ledgers to repair ─────────────────────────────────────────────────
# Tally ledger name  →  Odoo partner search term (name search)
REPAIR_MAP = {
    "USAM TECHNOLOGY SOLUTIONS PRIVATE LIMITED":          "USAM TECHNOLOGY SOLUTIONS",
    "CONSTELLATION AEROSPACE PRIVATE LIMITED":            "CONSTELLATION AEROSPACE",
    "IIIT DHARWAD RESEARCH PARK FOUNDATION, Suhruth Ujjni": "IIIT DHARWAD RESEARCH PARK",
    "Gleam Innovations, Akul Purushotham":                "Gleam Innovations",
    "ADSYNDICATE SERVICES PRIVATE LIMITED":               "ADSYNDICATE SERVICES",
}


def _build_kwargs(p: dict) -> dict:
    """Extract contact/tax kwargs from an Odoo partner record."""
    kwargs = {}
    addr = []
    if p.get("street"):  addr.append(p["street"])
    if p.get("street2"): addr.append(p["street2"])
    city_zip = " - ".join(filter(None, [p.get("city"), p.get("zip")]))
    if city_zip: addr.append(city_zip)
    if addr: kwargs["address_lines"] = addr

    state = p.get("state_id")
    if isinstance(state, (list, tuple)) and len(state) >= 2:
        kwargs["state"] = re.sub(r"\s*\([A-Z]{2}\)$", "", state[1]).strip()

    country = p.get("country_id")
    if isinstance(country, (list, tuple)) and len(country) >= 2:
        kwargs["country"] = country[1]

    if p.get("zip"):                    kwargs["pincode"]               = p["zip"]
    if p.get("phone"):                  kwargs["phone"]                 = p["phone"]
    if p.get("email"):                  kwargs["email"]                 = p["email"]
    if p.get("vat"):                    kwargs["gst_number"]            = p["vat"]
    if p.get("l10n_in_pan"):            kwargs["pan"]                   = p["l10n_in_pan"]
    if p.get("l10n_in_gst_treatment"):  kwargs["gst_registration_type"] = p["l10n_in_gst_treatment"]
    return kwargs


def main():
    parser = argparse.ArgumentParser(description="Repair bare Tally ledger details")
    parser.add_argument("--execute", action="store_true",
                        help="Actually ALTER ledgers (default: dry run)")
    args = parser.parse_args()

    odoo  = OdooConnector()
    tally = TallyConnector()
    group = "Sundry Debtors"

    log.info("=" * 65)
    log.info(f"{'LIVE ALTER' if args.execute else 'DRY RUN'} — Repair {len(REPAIR_MAP)} ledger(s)")
    log.info("=" * 65)

    for tally_name, search_term in REPAIR_MAP.items():
        log.info(f"\nLedger : {tally_name}")

        # Find partner in Odoo by name
        partners = odoo.search_read(
            "res.partner",
            [("name", "ilike", search_term)],
            ["id", "name", "vat", "state_id", "zip", "city"],
            limit=5,
        )
        if not partners:
            log.warning(f"  [SKIP] No Odoo partner found for search: '{search_term}'")
            continue

        # Use first match (most relevant by ilike)
        partner = partners[0]
        log.info(f"  Odoo  : [{partner['id']}] {partner['name']}")

        # Fetch full details
        details = odoo.get_partner_details(partner["id"])
        kwargs  = _build_kwargs(details)

        log.info(f"  GST   : {kwargs.get('gst_number') or '(none)'}")
        log.info(f"  State : {kwargs.get('state') or '(none)'}")
        log.info(f"  PIN   : {kwargs.get('pincode') or '(none)'}")
        log.info(f"  Phone : {kwargs.get('phone') or kwargs.get('mobile') or '(none)'}")
        log.info(f"  Email : {kwargs.get('email') or '(none)'}")

        if not args.execute:
            log.info("  [DRY RUN] Would ALTER this ledger in Tally")
            continue

        result = tally.create_ledger(tally_name, group, action="Alter", **kwargs)
        if result.startswith("OK"):
            log.info(f"  [OK] {result}")
        else:
            log.error(f"  [FAIL] {result}")

    if not args.execute:
        log.info("\n  Re-run with --execute to apply changes.")


if __name__ == "__main__":
    main()
