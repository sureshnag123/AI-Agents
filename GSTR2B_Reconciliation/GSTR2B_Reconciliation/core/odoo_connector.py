"""
Odoo XML-RPC Connector for GSTR2B Reconciliation.

Connects to an Odoo instance and fetches posted vendor bills (Purchase invoices)
for a given month, returning a DataFrame in the same schema as load_tally_data().

Odoo uses Python's built-in xmlrpc.client — no extra packages required.
"""
import xmlrpc.client
import logging
import ssl
from calendar import monthrange
from datetime import datetime

import pandas as pd

from core.utils import (
    clean_gstin, normalize_invoice_number, normalize_name, safe_float,
)

logger = logging.getLogger("gstr2b_reco")


class OdooConnector:
    """
    Thin wrapper around Odoo XML-RPC API (v2).

    Usage:
        conn = OdooConnector(url, db, username, api_key)
        conn.connect()          # authenticates; raises on failure
        df = conn.fetch_vendor_bills("2026-03")
    """

    def __init__(self, url: str, db: str, username: str, api_key: str):
        self.url = url.rstrip("/")
        self.db = db
        self.username = username
        self.api_key = api_key       # Odoo API key or password
        self.uid: int | None = None
        self._common = None
        self._models = None

    # ------------------------------------------------------------------
    # Connection
    # ------------------------------------------------------------------
    def connect(self) -> int:
        """
        Authenticate against the Odoo instance.
        Returns the user-id (uid) on success; raises ValueError on failure.
        """
        # Some self-hosted Odoo instances have unverified SSL — allow override
        ctx = ssl._create_unverified_context()

        self._common = xmlrpc.client.ServerProxy(
            f"{self.url}/xmlrpc/2/common",
            context=ctx,
            allow_none=True,
        )
        self._models = xmlrpc.client.ServerProxy(
            f"{self.url}/xmlrpc/2/object",
            context=ctx,
            allow_none=True,
        )

        version = self._common.version()
        logger.info(f"Odoo server version: {version.get('server_version', 'unknown')}")

        self.uid = self._common.authenticate(self.db, self.username, self.api_key, {})
        if not self.uid:
            raise ValueError(
                "Odoo authentication failed. "
                "Please check the URL, database name, username, and API key / password."
            )
        logger.info(f"Authenticated as uid={self.uid} on db='{self.db}'")
        return self.uid

    # ------------------------------------------------------------------
    # Low-level execute
    # ------------------------------------------------------------------
    def _execute(self, model: str, method: str, args: list, kwargs: dict | None = None) -> list:
        if self.uid is None:
            raise RuntimeError("Not connected. Call connect() first.")
        return self._models.execute_kw(
            self.db, self.uid, self.api_key,
            model, method, args, kwargs or {},
        )

    # ------------------------------------------------------------------
    # Fetch Vendor Bills
    # ------------------------------------------------------------------
    def fetch_vendor_bills(self, month: str) -> pd.DataFrame:
        """
        Fetch all posted vendor bills (and debit notes) for the given month.

        Parameters
        ----------
        month : str
            Period in "YYYY-MM" format, e.g. "2026-03".

        Returns
        -------
        pd.DataFrame
            Same column schema as load_tally_data():
            voucher_type, voucher_date, voucher_date_raw, voucher_number,
            party_name, party_name_norm, gstin, bill_number_raw, bill_number,
            taxable_value, igst, cgst, sgst, total_tax, total_value,
            narration, source
        """
        year, mon = map(int, month.split("-"))
        last_day = monthrange(year, mon)[1]
        date_from = f"{year}-{mon:02d}-01"
        date_to = f"{year}-{mon:02d}-{last_day:02d}"

        logger.info(f"Fetching Odoo vendor bills for {date_from} to {date_to}")

        # Search for vendor bills
        domain = [
            ["move_type", "in", ["in_invoice", "in_refund"]],
            ["invoice_date", ">=", date_from],
            ["invoice_date", "<=", date_to],
            ["state", "=", "posted"],
        ]

        inv_fields = [
            "id", "name", "ref", "move_type",
            "invoice_date", "partner_id",
            "amount_untaxed", "amount_tax", "amount_total",
        ]

        invoices = self._execute(
            "account.move", "search_read", [domain], {"fields": inv_fields, "order": "invoice_date asc"}
        )

        if not invoices:
            logger.warning(f"No vendor bills found in Odoo for {month}")
            return pd.DataFrame()

        logger.info(f"Found {len(invoices)} vendor bills in Odoo")

        # ------------------------------------------------------------------
        # Fetch partner details (name + VAT/GSTIN)
        # ------------------------------------------------------------------
        partner_ids = list({
            inv["partner_id"][0]
            for inv in invoices
            if inv.get("partner_id")
        })
        raw_partners = self._execute(
            "res.partner", "read",
            [partner_ids],
            {"fields": ["id", "name", "vat"]},
        )
        partner_map = {p["id"]: p for p in raw_partners}

        # ------------------------------------------------------------------
        # Fetch tax lines per invoice to split IGST / CGST / SGST
        # ------------------------------------------------------------------
        move_ids = [inv["id"] for inv in invoices]

        # account.move.line rows that are tax lines (tax_line_id is set)
        tax_line_fields = ["move_id", "tax_line_id", "name", "balance", "debit", "credit"]
        raw_tax_lines = self._execute(
            "account.move.line",
            "search_read",
            [[["move_id", "in", move_ids], ["tax_line_id", "!=", False]]],
            {"fields": tax_line_fields},
        )

        # Group by move_id
        tax_by_move: dict[int, list] = {}
        for tl in raw_tax_lines:
            mid = tl["move_id"][0]
            tax_by_move.setdefault(mid, []).append(tl)

        # ------------------------------------------------------------------
        # Build output rows
        # ------------------------------------------------------------------
        rows = []
        for inv in invoices:
            pid = inv["partner_id"][0] if inv.get("partner_id") else None
            partner = partner_map.get(pid, {})

            # Vendor reference (bill number) — prefer "ref" (vendor bill no.), fall back to "name"
            vendor_ref = (inv.get("ref") or "").strip()
            odoo_name = (inv.get("name") or "").strip()
            bill_number_raw = vendor_ref or odoo_name

            # GST tax breakdown
            igst = cgst = sgst = 0.0
            for tl in tax_by_move.get(inv["id"], []):
                # Tax name is in tl["name"] or tl["tax_line_id"][1]
                tax_name = ""
                if tl.get("name"):
                    tax_name = str(tl["name"]).upper()
                elif isinstance(tl.get("tax_line_id"), (list, tuple)) and len(tl["tax_line_id"]) > 1:
                    tax_name = str(tl["tax_line_id"][1]).upper()

                # For in_invoice (vendor bill), tax lines are typically debit (input credit)
                # Use abs(balance) to handle both debit and credit conventions
                amt = abs(safe_float(tl.get("balance", 0)))

                if "IGST" in tax_name:
                    igst += amt
                elif "CGST" in tax_name:
                    cgst += amt
                elif "SGST" in tax_name or "UTGST" in tax_name:
                    sgst += amt
                # If tax name has no GST tag, it's included in amount_tax already

            taxable_value = safe_float(inv.get("amount_untaxed", 0))
            amount_tax = safe_float(inv.get("amount_tax", 0))
            amount_total = safe_float(inv.get("amount_total", 0))

            # If we couldn't split into IGST/CGST/SGST (non-GST taxes, etc.),
            # fall back to total tax in IGST for matching purposes
            computed_tax = round(igst + cgst + sgst, 2)
            if computed_tax == 0 and amount_tax > 0:
                igst = amount_tax   # fallback: put all tax in IGST

            rows.append({
                "voucher_type": "Purchase" if inv.get("move_type") == "in_invoice" else "Debit Note",
                "voucher_date_raw": inv.get("invoice_date", ""),
                "voucher_date": _parse_odoo_date(inv.get("invoice_date")),
                "voucher_number": odoo_name,
                "party_name": (partner.get("name") or "").strip(),
                "gstin": clean_gstin(partner.get("vat") or ""),
                "bill_number_raw": bill_number_raw,
                "bill_number": normalize_invoice_number(bill_number_raw),
                "taxable_value": taxable_value,
                "igst": igst,
                "cgst": cgst,
                "sgst": sgst,
                "total_tax": amount_tax,
                "total_value": amount_total,
                "narration": vendor_ref,   # vendor ref doubles as narration context
                "source": "Odoo",
            })

        df = pd.DataFrame(rows)
        df["party_name_norm"] = df["party_name"].apply(normalize_name)

        logger.info(f"Odoo: {len(df)} vendor bills loaded for {month}")
        return df


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
def _parse_odoo_date(val) -> datetime | None:
    """Convert Odoo date string (YYYY-MM-DD) to datetime."""
    if not val:
        return None
    try:
        return datetime.strptime(str(val).strip()[:10], "%Y-%m-%d")
    except ValueError:
        return None
