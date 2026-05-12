#!/usr/bin/env python3
"""
Odoo XML-RPC Connector

Provides a reusable connection class for Odoo 19 using XML-RPC.
Handles authentication, model access, and common CRUD operations.

Usage:
    from odoo_connector import OdooConnector
    odoo = OdooConnector()
    invoices = odoo.search_read('account.move', [('state','=','posted')], ['name','partner_id'])
"""

import os
import xmlrpc.client
from dotenv import load_dotenv
from typing import Optional

load_dotenv()


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
        self.password = password or os.getenv("ODOO_PASSWORD", "")  # Can also be an API key

        if not all([self.url, self.db, self.username, self.password]):
            raise ValueError(
                "Missing Odoo connection details. Set ODOO_URL, ODOO_DB, "
                "ODOO_USERNAME, and ODOO_PASSWORD in your .env file."
            )

        # XML-RPC endpoints
        self._common = xmlrpc.client.ServerProxy(f"{self.url}/xmlrpc/2/common")
        self._models = xmlrpc.client.ServerProxy(f"{self.url}/xmlrpc/2/object")
        self._uid = None

    # ── Authentication ──────────────────────────────────────────────

    @property
    def uid(self) -> int:
        """Authenticate and cache the user ID."""
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
        """Return the Odoo server version info."""
        return self._common.version()

    # ── Generic CRUD helpers ────────────────────────────────────────

    def _execute(self, model: str, method: str, *args, **kwargs):
        """Low-level execute_kw wrapper."""
        return self._models.execute_kw(
            self.db, self.uid, self.password, model, method, list(args), kwargs
        )

    def search(self, model: str, domain: list, **kw) -> list[int]:
        """Return record IDs matching *domain*."""
        return self._execute(model, "search", domain, **kw)

    def read(self, model: str, ids: list[int], fields: list[str]) -> list[dict]:
        """Read specific fields from given IDs."""
        return self._execute(model, "read", ids, {"fields": fields})

    def search_read(
        self,
        model: str,
        domain: list,
        fields: list[str],
        limit: int = 0,
        order: str = "",
    ) -> list[dict]:
        """Search + read in one call."""
        kw = {"fields": fields}
        if limit:
            kw["limit"] = limit
        if order:
            kw["order"] = order
        return self._execute(model, "search_read", domain, **kw)

    def search_count(self, model: str, domain: list) -> int:
        """Return the count of records matching *domain*."""
        return self._execute(model, "search_count", domain)

    def create(self, model: str, values: dict) -> int:
        """Create a record and return its ID."""
        return self._execute(model, "create", values)

    def write(self, model: str, ids: list[int], values: dict) -> bool:
        """Update records."""
        return self._execute(model, "write", ids, values)

    def unlink(self, model: str, ids: list[int]) -> bool:
        """Delete records."""
        return self._execute(model, "unlink", ids)

    # ── Invoice-specific helpers ────────────────────────────────────

    def get_overdue_invoices(self, days_overdue: int = 0) -> list[dict]:
        """
        Fetch posted customer invoices whose due date is at or past today
        minus *days_overdue* buffer.

        Returns a list of dicts with key invoice fields.
        """
        from datetime import date, timedelta

        cutoff = date.today() - timedelta(days=days_overdue)

        domain = [
            ("move_type", "=", "out_invoice"),
            ("state", "=", "posted"),
            ("payment_state", "in", ["not_paid", "partial"]),
            ("invoice_date_due", "<=", str(cutoff)),
        ]

        fields = [
            "id",
            "name",                    # INV/2026/0001
            "partner_id",              # [id, name]
            "invoice_date",
            "invoice_date_due",
            "amount_total",
            "amount_residual",         # remaining amount
            "currency_id",
            "payment_state",
            "invoice_payment_term_id", # payment term used
            "user_id",                 # salesperson
        ]

        return self.search_read("account.move", domain, fields, order="invoice_date_due asc")

    def get_upcoming_due_invoices(self, days_ahead: int = 7) -> list[dict]:
        """
        Fetch posted customer invoices due within the next *days_ahead* days.
        Useful for sending pre-due reminders.
        """
        from datetime import date, timedelta

        today = date.today()
        future = today + timedelta(days=days_ahead)

        domain = [
            ("move_type", "=", "out_invoice"),
            ("state", "=", "posted"),
            ("payment_state", "in", ["not_paid", "partial"]),
            ("invoice_date_due", ">=", str(today)),
            ("invoice_date_due", "<=", str(future)),
        ]

        fields = [
            "id",
            "name",
            "partner_id",
            "invoice_date",
            "invoice_date_due",
            "amount_total",
            "amount_residual",
            "currency_id",
            "payment_state",
            "invoice_payment_term_id",
            "user_id",
        ]

        return self.search_read("account.move", domain, fields, order="invoice_date_due asc")

    def get_partner_email(self, partner_id: int) -> dict:
        """Return partner contact details (name, email, phone)."""
        records = self.read(
            "res.partner",
            [partner_id],
            ["name", "email", "phone", "mobile", "lang"],
        )
        return records[0] if records else {}

    def get_payment_term_details(self, term_id: int) -> dict:
        """Return payment term name and line details."""
        records = self.read(
            "account.payment.term",
            [term_id],
            ["name", "note", "line_ids"],
        )
        return records[0] if records else {}

    def log_reminder_on_invoice(self, invoice_id: int, message: str) -> int:
        """Post an internal note (chatter message) on the invoice."""
        return self._execute(
            "account.move",
            "message_post",
            [invoice_id],
            body=message,
            message_type="comment",
            subtype_xmlid="mail.mt_note",
        )


# ── Quick CLI test ──────────────────────────────────────────────────

if __name__ == "__main__":
    import json

    odoo = OdooConnector()
    print(f"Connected to Odoo {odoo.version().get('server_version', '?')}")
    print(f"Authenticated as UID={odoo.uid}\n")

    overdue = odoo.get_overdue_invoices()
    print(f"Overdue invoices: {len(overdue)}")
    for inv in overdue[:5]:
        print(
            f"  {inv['name']}  |  {inv['partner_id'][1]}  |  "
            f"Due: {inv['invoice_date_due']}  |  "
            f"Remaining: {inv['amount_residual']} {inv['currency_id'][1]}"
        )

    upcoming = odoo.get_upcoming_due_invoices()
    print(f"\nUpcoming due (next 7 days): {len(upcoming)}")
    for inv in upcoming[:5]:
        print(
            f"  {inv['name']}  |  {inv['partner_id'][1]}  |  "
            f"Due: {inv['invoice_date_due']}  |  "
            f"Remaining: {inv['amount_residual']} {inv['currency_id'][1]}"
        )
