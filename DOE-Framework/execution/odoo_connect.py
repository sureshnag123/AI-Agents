"""
odoo_connect.py — Odoo XML-RPC client utility
Base connection and CRUD helpers used by all Odoo agent scripts.

Usage (as a library):
    from odoo_connect import OdooClient
    client = OdooClient()
    records = client.search_read('sale.order', [('state', '=', 'sale')], ['name', 'date_order'])

Usage (as CLI — connection test):
    python execution/odoo_connect.py
"""

import os
import sys
import json
import xmlrpc.client
from datetime import datetime, timedelta
from dotenv import load_dotenv

load_dotenv()


class OdooClient:
    """Authenticated Odoo XML-RPC client."""

    def __init__(self):
        self.url = os.getenv("ODOO_URL", "").rstrip("/")
        self.db = os.getenv("ODOO_DB", "")
        self.username = os.getenv("ODOO_USERNAME", "")
        self.api_key = os.getenv("ODOO_API_KEY", "")

        if not all([self.url, self.db, self.username, self.api_key]):
            missing = [k for k, v in {
                "ODOO_URL": self.url,
                "ODOO_DB": self.db,
                "ODOO_USERNAME": self.username,
                "ODOO_API_KEY": self.api_key,
            }.items() if not v]
            raise ValueError(f"Missing required env vars: {', '.join(missing)}")

        self._common = xmlrpc.client.ServerProxy(f"{self.url}/xmlrpc/2/common")
        self._models = xmlrpc.client.ServerProxy(f"{self.url}/xmlrpc/2/object")
        self._uid = None

    @property
    def uid(self):
        if self._uid is None:
            self._uid = self._common.authenticate(self.db, self.username, self.api_key, {})
            if not self._uid:
                raise ConnectionError(
                    "Odoo authentication failed. Check ODOO_URL, ODOO_DB, ODOO_USERNAME, ODOO_API_KEY."
                )
        return self._uid

    def execute(self, model, method, *args, **kwargs):
        """Raw execute_kw call."""
        return self._models.execute_kw(
            self.db, self.uid, self.api_key,
            model, method, list(args), kwargs
        )

    def search(self, model, domain, **kwargs):
        """Return list of IDs matching domain."""
        return self.execute(model, "search", domain, **kwargs)

    def read(self, model, ids, fields=None, **kwargs):
        """Read records by IDs."""
        return self.execute(model, "read", ids, fields=fields or [], **kwargs)

    def search_read(self, model, domain, fields=None, limit=500, order=None, offset=0):
        """Search and read in one call."""
        kwargs = {"limit": limit, "offset": offset}
        if fields:
            kwargs["fields"] = fields
        if order:
            kwargs["order"] = order
        return self.execute(model, "search_read", domain, **kwargs)

    def count(self, model, domain):
        """Count records matching domain."""
        return self.execute(model, "search_count", domain)

    def get_installed_modules(self):
        """Return list of installed module names."""
        modules = self.search_read(
            "ir.module.module",
            [("state", "=", "installed")],
            fields=["name", "shortdesc"],
            limit=1000
        )
        return {m["name"]: m["shortdesc"] for m in modules}

    def date_domain(self, field, days_back):
        """Generate a date filter domain for the last N days."""
        start = (datetime.now() - timedelta(days=days_back)).strftime("%Y-%m-%d 00:00:00")
        return [(field, ">=", start)]

    def server_info(self):
        """Return Odoo server version info."""
        return self._common.version()


def test_connection():
    """CLI: test Odoo connection and print basic info."""
    print("Testing Odoo connection...")
    try:
        client = OdooClient()
        uid = client.uid
        info = client.server_info()
        print(f"  Connected as UID: {uid}")
        print(f"  Odoo version: {info.get('server_version', 'unknown')}")

        # Quick record counts
        so_count = client.count("sale.order", [("state", "in", ["sale", "done"])])
        po_count = client.count("purchase.order", [("state", "in", ["purchase", "done"])])
        mo_count = client.count("mrp.production", [("state", "not in", ["cancel"])])

        print(f"\n  Confirmed Sales Orders : {so_count}")
        print(f"  Purchase Orders        : {po_count}")
        print(f"  Manufacturing Orders   : {mo_count}")
        print("\nConnection OK.")
        return True
    except Exception as e:
        print(f"\nERROR: {e}")
        print("\nTroubleshooting:")
        print("  1. Verify ODOO_URL (include https://, no trailing slash)")
        print("  2. Verify ODOO_DB (exact database name)")
        print("  3. Verify ODOO_USERNAME (login email)")
        print("  4. Generate API key: Odoo → Settings → Technical → API Keys")
        return False


if __name__ == "__main__":
    ok = test_connection()
    sys.exit(0 if ok else 1)
