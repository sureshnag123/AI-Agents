#!/usr/bin/env python3
"""
Odoo XML-RPC Connector for Odoo-Tally Sync

Provides a reusable connection class for Odoo 19 using XML-RPC.
Handles authentication and provides helpers for fetching invoices,
payments, and inventory movements for Tally sync.

Usage:
    from odoo_connector import OdooConnector
    odoo = OdooConnector()
    invoices = odoo.get_purchase_invoices(from_date="2026-02-01")
    
    # CLI test
    python odoo_connector.py --test
    python odoo_connector.py --list-accounts
"""

import os
import sys
import json
import argparse
import xmlrpc.client
from datetime import date, timedelta, datetime
from typing import Optional
from pathlib import Path

from dotenv import load_dotenv

SCRIPT_DIR = Path(__file__).parent.resolve()
PROJECT_ROOT = SCRIPT_DIR.parent
TMP_DIR = PROJECT_ROOT / ".tmp"

load_dotenv(PROJECT_ROOT / ".env")


class OdooConnector:
    """XML-RPC connector for Odoo 19."""

    def __init__(
        self,
        url: Optional[str] = None,
        db: Optional[str] = None,
        username: Optional[str] = None,
        password: Optional[str] = None,
    ):
        self.url = (url or os.getenv("ODOO_URL", "")).rstrip("/")
        self.db = db or os.getenv("ODOO_DB", "")
        self.username = username or os.getenv("ODOO_USERNAME", "")
        self.password = password or os.getenv("ODOO_PASSWORD", "")

        if not all([self.url, self.db, self.username, self.password]):
            raise ValueError(
                "Missing Odoo connection details. Set ODOO_URL, ODOO_DB, "
                "ODOO_USERNAME, and ODOO_PASSWORD in your .env file."
            )

        self._common = xmlrpc.client.ServerProxy(f"{self.url}/xmlrpc/2/common")
        self._models = xmlrpc.client.ServerProxy(f"{self.url}/xmlrpc/2/object")
        self._uid = None

    # ── Authentication ──────────────────────────────────────────────

    @property
    def uid(self) -> int:
        if self._uid is None:
            self._uid = self._common.authenticate(
                self.db, self.username, self.password, {}
            )
            if not self._uid:
                raise ConnectionError(
                    f"Odoo authentication failed for user '{self.username}' "
                    f"on database '{self.db}' at {self.url}"
                )
        return self._uid

    def version(self) -> dict:
        return self._common.version()

    # ── Generic CRUD ────────────────────────────────────────────────

    def _execute(self, model: str, method: str, *args, **kwargs):
        return self._models.execute_kw(
            self.db, self.uid, self.password, model, method, list(args), kwargs
        )

    def search(self, model: str, domain: list, **kw) -> list:
        return self._execute(model, "search", domain, **kw)

    def read(self, model: str, ids: list, fields: list) -> list:
        return self._execute(model, "read", ids, fields=fields)

    def search_read(self, model: str, domain: list, fields: list,
                    limit: int = 0, order: str = "", offset: int = 0) -> list:
        kw = {"fields": fields}
        if limit:
            kw["limit"] = limit
        if order:
            kw["order"] = order
        if offset:
            kw["offset"] = offset
        return self._execute(model, "search_read", domain, **kw)

    def search_count(self, model: str, domain: list) -> int:
        return self._execute(model, "search_count", domain)

    # ── Invoice Helpers (account.move) ──────────────────────────────

    INVOICE_FIELDS = [
        "id", "name", "ref", "move_type", "state",
        "partner_id", "invoice_date", "date", "invoice_date_due",
        "amount_total", "amount_untaxed", "amount_tax", "amount_residual",
        "currency_id", "payment_state",
        "invoice_payment_term_id", "journal_id",
        "invoice_line_ids", "line_ids",
        "create_date", "write_date",
    ]

    def _get_invoices(self, move_type: str, from_date: Optional[str] = None,
                      to_date: Optional[str] = None,
                      since_write_date: Optional[str] = None,
                      limit: int = 0) -> list:
        """Fetch posted invoices of a given type."""
        domain = [
            ("move_type", "=", move_type),
            ("state", "=", "posted"),
        ]
        if from_date:
            domain.append(("invoice_date", ">=", from_date))
        if to_date:
            domain.append(("invoice_date", "<=", to_date))
        if since_write_date:
            domain.append(("write_date", ">=", since_write_date))

        return self.search_read(
            "account.move", domain, self.INVOICE_FIELDS,
            limit=limit, order="invoice_date asc"
        )

    def get_purchase_invoices(self, **kwargs) -> list:
        """Fetch posted vendor/purchase invoices (Bills)."""
        return self._get_invoices("in_invoice", **kwargs)

    def get_purchase_returns(self, **kwargs) -> list:
        """Fetch posted vendor debit notes (purchase returns)."""
        return self._get_invoices("in_refund", **kwargs)

    def get_sales_invoices(self, **kwargs) -> list:
        """Fetch posted customer invoices."""
        return self._get_invoices("out_invoice", **kwargs)

    def get_sales_returns(self, **kwargs) -> list:
        """Fetch posted customer credit notes (sales returns)."""
        return self._get_invoices("out_refund", **kwargs)

    # ── Invoice Line Details ────────────────────────────────────────

    def get_invoice_lines(self, invoice_id: int) -> list:
        """Fetch line items for a specific invoice."""
        return self.search_read(
            "account.move.line",
            [("move_id", "=", invoice_id)],
            [
                "id", "name", "account_id", "account_type",
                "partner_id",
                "product_id", "product_uom_id",
                "quantity", "price_unit", "price_subtotal", "price_total",
                "discount", "tax_ids", "tax_line_id",
                "debit", "credit", "balance",
                "amount_currency",
            ],
        )

    def get_tax_details(self, tax_ids: list) -> list:
        """Fetch tax configuration details."""
        if not tax_ids:
            return []
        return self.read(
            "account.tax",
            tax_ids,
            ["id", "name", "amount", "amount_type", "type_tax_use",
             "tax_group_id", "description"],
        )

    # ── Payment Helpers (account.payment) ───────────────────────────

    PAYMENT_FIELDS = [
        "id", "name", "ref", "payment_type", "partner_type",
        "partner_id", "amount", "currency_id",
        "journal_id", "payment_method_id",
        "date", "state",
        "reconciled_invoice_ids",
        "create_date", "write_date",
    ]

    def get_outbound_payments(self, from_date: Optional[str] = None,
                              to_date: Optional[str] = None,
                              since_write_date: Optional[str] = None) -> list:
        """Fetch posted outbound payments (vendor payments)."""
        domain = [
            ("payment_type", "=", "outbound"),
            ("state", "=", "posted"),
        ]
        if from_date:
            domain.append(("date", ">=", from_date))
        if to_date:
            domain.append(("date", "<=", to_date))
        if since_write_date:
            domain.append(("write_date", ">=", since_write_date))

        return self.search_read(
            "account.payment", domain, self.PAYMENT_FIELDS,
            order="date asc"
        )

    def get_inbound_payments(self, from_date: Optional[str] = None,
                             to_date: Optional[str] = None,
                             since_write_date: Optional[str] = None) -> list:
        """Fetch posted inbound payments (customer receipts)."""
        domain = [
            ("payment_type", "=", "inbound"),
            ("state", "=", "posted"),
        ]
        if from_date:
            domain.append(("date", ">=", from_date))
        if to_date:
            domain.append(("date", "<=", to_date))
        if since_write_date:
            domain.append(("write_date", ">=", since_write_date))

        return self.search_read(
            "account.payment", domain, self.PAYMENT_FIELDS,
            order="date asc"
        )

    # ── Account / Ledger Helpers ────────────────────────────────────

    def get_chart_of_accounts(self) -> list:
        """Fetch all accounts from Odoo's chart of accounts."""
        return self.search_read(
            "account.account",
            [],
            ["id", "code", "name", "account_type", "reconcile",
             "tax_ids", "group_id"],
            order="code asc"
        )

    def get_partner_details(self, partner_id: int) -> dict:
        """Fetch partner name, GST, PAN, address, and Indian localization fields."""
        base_fields = [
            "name", "vat", "street", "street2", "city", "state_id",
            "zip", "country_id", "email", "phone",
            "is_company", "company_type", "property_account_receivable_id",
            "property_account_payable_id",
        ]
        # Indian localization fields (present only when l10n_in is installed)
        india_fields = ["l10n_in_gst_treatment", "l10n_in_pan"]
        try:
            records = self.read("res.partner", [partner_id], base_fields + india_fields)
        except Exception:
            # Fallback: fetch without India-specific fields if the module is absent
            records = self.read("res.partner", [partner_id], base_fields)
        return records[0] if records else {}

    def get_all_partners(self, supplier: bool = False, customer: bool = False) -> list:
        """Fetch all partners (optionally filtered by type)."""
        domain = []
        if supplier:
            domain.append(("supplier_rank", ">", 0))
        if customer:
            domain.append(("customer_rank", ">", 0))
        return self.search_read(
            "res.partner", domain,
            ["id", "name", "vat", "city", "state_id", "is_company"],
            order="name asc"
        )

    # ── Inventory / Stock Helpers ───────────────────────────────────

    def get_stock_moves(self, from_date: Optional[str] = None,
                        to_date: Optional[str] = None) -> list:
        """Fetch completed stock moves (for inventory journal entries)."""
        domain = [("state", "=", "done")]
        if from_date:
            domain.append(("date", ">=", from_date))
        if to_date:
            domain.append(("date", "<=", to_date))

        return self.search_read(
            "stock.move", domain,
            ["id", "name", "reference", "product_id", "product_uom_qty",
             "product_uom", "location_id", "location_dest_id",
             "origin", "date", "picking_id", "state",
             "price_unit", "create_date", "write_date"],
            order="date asc"
        )

    # ── Journal Helpers ─────────────────────────────────────────────

    def get_journals(self) -> list:
        """Fetch all accounting journals."""
        return self.search_read(
            "account.journal", [],
            ["id", "name", "code", "type", "default_account_id",
             "currency_id"],
            order="code asc"
        )


# ── CLI Entry Point ─────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description="Odoo 19 XML-RPC Connector")
    parser.add_argument("--test", action="store_true", help="Test Odoo connection")
    parser.add_argument("--list-accounts", action="store_true",
                        help="List all accounts from chart of accounts")
    parser.add_argument("--list-partners", action="store_true",
                        help="List all partners")
    parser.add_argument("--list-journals", action="store_true",
                        help="List all journals")
    parser.add_argument("--purchase-invoices", action="store_true",
                        help="List recent purchase invoices")
    parser.add_argument("--sales-invoices", action="store_true",
                        help="List recent sales invoices")
    parser.add_argument("--from-date", type=str, default=None,
                        help="Start date (YYYY-MM-DD)")
    parser.add_argument("--to-date", type=str, default=None,
                        help="End date (YYYY-MM-DD)")
    parser.add_argument("--limit", type=int, default=20,
                        help="Max records to display (default: 20)")

    args = parser.parse_args()
    TMP_DIR.mkdir(parents=True, exist_ok=True)

    try:
        odoo = OdooConnector()
    except ValueError as e:
        print(f"ERROR: {e}")
        sys.exit(1)

    if args.test:
        try:
            ver = odoo.version()
            print(f"✓ Connected to Odoo {ver.get('server_version', '?')}")
            print(f"  URL:      {odoo.url}")
            print(f"  Database: {odoo.db}")
            print(f"  User:     {odoo.username}")
            print(f"  UID:      {odoo.uid}")

            # Quick counts
            pi = odoo.search_count("account.move",
                                   [("move_type", "=", "in_invoice"), ("state", "=", "posted")])
            si = odoo.search_count("account.move",
                                   [("move_type", "=", "out_invoice"), ("state", "=", "posted")])
            print(f"\n  Posted Purchase Invoices: {pi}")
            print(f"  Posted Sales Invoices:    {si}")
        except Exception as e:
            print(f"✗ Connection failed: {e}")
            sys.exit(1)

    elif args.list_accounts:
        accounts = odoo.get_chart_of_accounts()
        print(f"\n{'Code':<10} {'Name':<50} {'Type':<25}")
        print("─" * 85)
        for acc in accounts[:args.limit]:
            print(f"{acc.get('code',''):<10} {acc['name']:<50} {acc.get('account_type',''):<25}")
        if len(accounts) > args.limit:
            print(f"\n... and {len(accounts) - args.limit} more accounts")
        # Save full list to tmp
        out = TMP_DIR / "odoo_accounts.json"
        with open(out, "w", encoding="utf-8") as f:
            json.dump(accounts, f, indent=2, default=str)
        print(f"\nFull list saved to {out}")

    elif args.list_partners:
        partners = odoo.get_all_partners()
        print(f"\n{'ID':<8} {'Name':<50} {'GSTIN':<20} {'City':<20}")
        print("─" * 98)
        for p in partners[:args.limit]:
            vat = p.get("vat") or ""
            city = p.get("city") or ""
            print(f"{p['id']:<8} {p['name']:<50} {vat:<20} {city:<20}")
        if len(partners) > args.limit:
            print(f"\n... and {len(partners) - args.limit} more partners")
        out = TMP_DIR / "odoo_partners.json"
        with open(out, "w", encoding="utf-8") as f:
            json.dump(partners, f, indent=2, default=str)
        print(f"\nFull list saved to {out}")

    elif args.list_journals:
        journals = odoo.get_journals()
        print(f"\n{'Code':<8} {'Name':<30} {'Type':<15}")
        print("─" * 53)
        for j in journals:
            print(f"{j.get('code',''):<8} {j['name']:<30} {j.get('type',''):<15}")

    elif args.purchase_invoices:
        invoices = odoo.get_purchase_invoices(
            from_date=args.from_date, to_date=args.to_date, limit=args.limit
        )
        print(f"\n{'Invoice #':<20} {'Vendor':<35} {'Date':<12} {'Total':>12} {'Status':<10}")
        print("─" * 89)
        for inv in invoices:
            partner = inv["partner_id"][1] if inv.get("partner_id") else "N/A"
            print(f"{inv['name']:<20} {partner:<35} {inv.get('invoice_date',''):<12} "
                  f"{inv['amount_total']:>12,.2f} {inv.get('payment_state',''):<10}")

    elif args.sales_invoices:
        invoices = odoo.get_sales_invoices(
            from_date=args.from_date, to_date=args.to_date, limit=args.limit
        )
        print(f"\n{'Invoice #':<20} {'Customer':<35} {'Date':<12} {'Total':>12} {'Status':<10}")
        print("─" * 89)
        for inv in invoices:
            partner = inv["partner_id"][1] if inv.get("partner_id") else "N/A"
            print(f"{inv['name']:<20} {partner:<35} {inv.get('invoice_date',''):<12} "
                  f"{inv['amount_total']:>12,.2f} {inv.get('payment_state',''):<10}")

    else:
        parser.print_help()


if __name__ == "__main__":
    main()
