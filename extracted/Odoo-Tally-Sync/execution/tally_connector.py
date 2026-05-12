#!/usr/bin/env python3
"""
Tally Prime XML Connector

Connects to Tally Prime's XML Server (HTTP on port 9000) to:
- Import vouchers (Purchase, Sales, Payment, Receipt, Journal, etc.)
- Export ledger lists, voucher data, and company info
- Create ledgers / stock items on-the-fly

Tally Prime must be running with XML Server enabled (F12 → Connectivity).

Usage:
    from tally_connector import TallyConnector
    tally = TallyConnector()
    tally.test_connection()
    ledgers = tally.get_all_ledgers()

    # CLI
    python tally_connector.py --test
    python tally_connector.py --list-ledgers
"""

import os
import re
import sys
import json
import argparse
import xml.etree.ElementTree as ET
from xml.sax.saxutils import escape as xml_escape
from pathlib import Path
from typing import Optional

import requests
from dotenv import load_dotenv

SCRIPT_DIR = Path(__file__).parent.resolve()
PROJECT_ROOT = SCRIPT_DIR.parent
TMP_DIR = PROJECT_ROOT / ".tmp"

load_dotenv(PROJECT_ROOT / ".env")


class TallyConnector:
    """HTTP/XML connector for Tally Prime."""

    def __init__(
        self,
        host: Optional[str] = None,
        port: Optional[int] = None,
        company: Optional[str] = None,
    ):
        self.host = host or os.getenv("TALLY_HOST", "localhost")
        self.port = int(port or os.getenv("TALLY_PORT", "9000"))
        self.company = company or os.getenv("TALLY_COMPANY_NAME", "")
        self.base_url = f"http://{self.host}:{self.port}"

    # ── Low-level XML Request ───────────────────────────────────────

    def _send_xml(self, xml_payload: str, timeout: int = 30) -> str:
        """Send raw XML to Tally and return the response body."""
        try:
            resp = requests.post(
                self.base_url,
                data=xml_payload.encode("utf-8"),
                headers={"Content-Type": "text/xml; charset=utf-8"},
                timeout=timeout,
            )
            resp.raise_for_status()
            return resp.text
        except requests.ConnectionError:
            raise ConnectionError(
                f"Cannot connect to Tally at {self.base_url}. "
                "Ensure Tally Prime is running with XML Server enabled "
                "(F12 → Connectivity → Tally Prime Server = Yes)."
            )
        except requests.Timeout:
            raise TimeoutError(
                f"Tally did not respond within {timeout}s. "
                "Try increasing timeout or check Tally load."
            )

    @staticmethod
    def _sanitize_xml(xml_text: str) -> str:
        """Remove invalid XML character references and namespace prefixes that Tally emits."""
        # Remove invalid XML numeric character references (control chars 0x00-0x08, 0x0B, 0x0C, 0x0E-0x1F)
        xml_text = re.sub(r'&#x([0-8BbCcEeFf]);', '', xml_text)
        xml_text = re.sub(r'&#x1[0-9A-Fa-f];', '', xml_text)
        xml_text = re.sub(r'&#([0-8]|1[0-1]|1[4-9]|2[0-9]|3[01]);', '', xml_text)
        # Remove raw control characters
        xml_text = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f]', '', xml_text)
        # Strip XML namespace prefixes (e.g. <UDF:FIELD> → <UDF_FIELD>)
        xml_text = re.sub(r'<(/?)(\w+):', r'<\1\2_', xml_text)
        xml_text = re.sub(r'\s+xmlns:\w+="[^"]*"', '', xml_text)
        return xml_text

    def _parse_response(self, xml_text: str) -> ET.Element:
        """Parse Tally XML response into an ElementTree Element."""
        xml_text = self._sanitize_xml(xml_text)
        try:
            return ET.fromstring(xml_text)
        except ET.ParseError as e:
            raise ValueError(f"Invalid XML response from Tally: {e}\n{xml_text[:500]}")

    # ── Connection Test ─────────────────────────────────────────────

    def test_connection(self) -> dict:
        """Test connectivity and return Tally license/company info."""
        xml = """<ENVELOPE>
  <HEADER>
    <VERSION>1</VERSION>
    <TALLYREQUEST>Export</TALLYREQUEST>
    <TYPE>Collection</TYPE>
    <ID>CompanyList</ID>
  </HEADER>
  <BODY>
    <DESC>
      <TDL>
        <TDLMESSAGE>
          <COLLECTION NAME="CompanyList" ISMODIFY="No">
            <TYPE>Company</TYPE>
            <FETCH>NAME, BASICCOMPANYFORMALNAME</FETCH>
          </COLLECTION>
        </TDLMESSAGE>
      </TDL>
    </DESC>
  </BODY>
</ENVELOPE>"""

        resp_text = self._send_xml(xml)
        root = self._parse_response(resp_text)

        companies = []
        for comp in root.iter("COMPANY"):
            name = comp.get("NAME") or comp.findtext("NAME", "")
            if name:
                companies.append(name.strip())

        return {
            "status": "connected",
            "url": self.base_url,
            "companies": companies,
            "target_company": self.company,
        }

    # ── Ledger Operations ───────────────────────────────────────────

    def get_all_ledgers(self) -> list:
        """Fetch all ledger names and their parent groups from Tally."""
        xml = f"""<ENVELOPE>
  <HEADER>
    <VERSION>1</VERSION>
    <TALLYREQUEST>Export</TALLYREQUEST>
    <TYPE>Collection</TYPE>
    <ID>Ledger Collection</ID>
  </HEADER>
  <BODY>
    <DESC>
      <STATICVARIABLES>
        <SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
        <SVCURRENTCOMPANY>{self.company}</SVCURRENTCOMPANY>
      </STATICVARIABLES>
      <TDL>
        <TDLMESSAGE>
          <COLLECTION NAME="Ledger Collection" ISMODIFY="No">
            <TYPE>Ledger</TYPE>
            <FETCH>NAME, PARENT, OPENINGBALANCE</FETCH>
          </COLLECTION>
        </TDLMESSAGE>
      </TDL>
    </DESC>
  </BODY>
</ENVELOPE>"""

        resp_text = self._send_xml(xml)
        root = self._parse_response(resp_text)

        ledgers = []
        for ldg in root.iter("LEDGER"):
            # NAME can be an attribute or nested in LANGUAGENAME.LIST
            name = (ldg.get("NAME") or ldg.findtext(".//NAME", "")).strip()
            parent = ldg.findtext("PARENT", "").strip()
            opening = ldg.findtext("OPENINGBALANCE", "0").strip()
            if name:
                ledgers.append({
                    "name": name,
                    "group": parent,
                    "opening_balance": opening,
                })
        return ledgers

    def get_all_groups(self) -> list:
        """Fetch all account groups from Tally."""
        xml = f"""<ENVELOPE>
  <HEADER>
    <VERSION>1</VERSION>
    <TALLYREQUEST>Export</TALLYREQUEST>
    <TYPE>Collection</TYPE>
    <ID>Group Collection</ID>
  </HEADER>
  <BODY>
    <DESC>
      <STATICVARIABLES>
        <SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
        <SVCURRENTCOMPANY>{self.company}</SVCURRENTCOMPANY>
      </STATICVARIABLES>
      <TDL>
        <TDLMESSAGE>
          <COLLECTION NAME="Group Collection" ISMODIFY="No">
            <TYPE>Group</TYPE>
            <FETCH>NAME, PARENT</FETCH>
          </COLLECTION>
        </TDLMESSAGE>
      </TDL>
    </DESC>
  </BODY>
</ENVELOPE>"""

        resp_text = self._send_xml(xml)
        root = self._parse_response(resp_text)

        groups = []
        for grp in root.iter("GROUP"):
            name = (grp.get("NAME") or grp.findtext(".//NAME", "")).strip()
            parent = grp.findtext("PARENT", "").strip()
            if name:
                groups.append({"name": name, "parent": parent})
        return groups

    def create_ledger(self, name: str, group: str, opening_balance: float = 0,
                      gst_number: str = "", address_lines: list = None,
                      state: str = "", pincode: str = "", country: str = "",
                      phone: str = "", mobile: str = "", email: str = "",
                      pan: str = "", gst_registration_type: str = "",
                      action: str = "Create") -> str:
        """Create or alter a ledger in Tally with full contact and tax details."""

        # ── GST fields ──────────────────────────────────────────────
        gst_xml = ""
        if gst_number:
            gst_xml = (
                f"<PARTYGSTIN>{xml_escape(gst_number)}</PARTYGSTIN>\n"
                f"          <GSTIN>{xml_escape(gst_number)}</GSTIN>"
            )

        # Map Odoo GST treatment → Tally registration type label
        _gst_type_map = {
            "regular": "Regular",
            "composition": "Composition",
            "unregistered": "Unregistered",
            "consumer": "Consumer",
            "overseas": "Overseas",
            "special_economic_zone": "Special Economic Zone",
            "deemed_export": "Deemed Export",
        }
        tally_gst_type = _gst_type_map.get(
            (gst_registration_type or "").lower(), gst_registration_type
        )
        gst_type_xml = ""
        if tally_gst_type:
            gst_type_xml = f"<GSTREGISTRATIONTYPE>{xml_escape(tally_gst_type)}</GSTREGISTRATIONTYPE>"

        # ── Multi-line address ───────────────────────────────────────
        addr_xml = ""
        lines = [l for l in (address_lines or []) if l and l.strip()]
        if lines:
            inner = "\n".join(f"          <ADDRESS>{xml_escape(l.strip())}</ADDRESS>" for l in lines)
            addr_xml = f"<ADDRESS.LIST TYPE=\"String\">\n{inner}\n          </ADDRESS.LIST>"

        # ── State / Country / Pincode ────────────────────────────────
        state_xml   = f"<STATENAME>{xml_escape(state)}</STATENAME>" if state else ""
        country_xml = f"<COUNTRYNAME>{xml_escape(country)}</COUNTRYNAME>" if country else ""
        pin_xml     = f"<PINCODE>{xml_escape(pincode)}</PINCODE>" if pincode else ""

        # ── Contact details ──────────────────────────────────────────
        phone_xml  = f"<LEDGERPHONE>{xml_escape(phone)}</LEDGERPHONE>" if phone else ""
        mobile_xml = f"<LEDGERMOBILE>{xml_escape(mobile)}</LEDGERMOBILE>" if mobile else ""
        email_xml  = f"<LEDGEREMAIL>{xml_escape(email)}</LEDGEREMAIL>" if email else ""

        # ── PAN ──────────────────────────────────────────────────────
        pan_xml = f"<INCOMETAXNUMBER>{xml_escape(pan)}</INCOMETAXNUMBER>" if pan else ""

        esc_name    = xml_escape(name)
        esc_group   = xml_escape(group)
        esc_action  = xml_escape(action)
        mailing_xml = f"<MAILINGNAME>{esc_name}</MAILINGNAME>"

        xml = f"""<ENVELOPE>
  <HEADER>
    <VERSION>1</VERSION>
    <TALLYREQUEST>Import</TALLYREQUEST>
    <TYPE>Data</TYPE>
    <ID>All Masters</ID>
  </HEADER>
  <BODY>
    <DESC>
      <STATICVARIABLES>
        <SVCURRENTCOMPANY>{self.company}</SVCURRENTCOMPANY>
      </STATICVARIABLES>
    </DESC>
    <DATA>
      <TALLYMESSAGE>
        <LEDGER NAME="{esc_name}" ACTION="{esc_action}">
          <NAME>{esc_name}</NAME>
          <PARENT>{esc_group}</PARENT>
          {mailing_xml}
          <OPENINGBALANCE>{opening_balance}</OPENINGBALANCE>
          {gst_xml}
          {gst_type_xml}
          {pan_xml}
          {addr_xml}
          {state_xml}
          {country_xml}
          {pin_xml}
          {phone_xml}
          {mobile_xml}
          {email_xml}
        </LEDGER>
      </TALLYMESSAGE>
    </DATA>
  </BODY>
</ENVELOPE>"""

        resp_text = self._send_xml(xml)
        root = self._parse_response(resp_text)

        # Check for errors
        created = root.findtext(".//CREATED", "0")
        altered = root.findtext(".//ALTERED", "0")
        errors  = root.findtext(".//ERRORS", "0")
        if int(errors) > 0:
            error_msg = root.findtext(".//LINEERROR", "Unknown error")
            return f"ERROR: {error_msg}"
        if action == "Alter":
            return f"OK: Ledger '{name}' updated (altered={altered})"
        return f"OK: Ledger '{name}' created under '{group}' (created={created})"

    def ledger_exists(self, name: str) -> bool:
        """Check if a ledger exists in Tally."""
        ledgers = self.get_all_ledgers()
        names_lower = {l["name"].lower() for l in ledgers}
        return name.lower() in names_lower

    # ── Voucher Import ──────────────────────────────────────────────

    def import_voucher_xml(self, voucher_xml: str, timeout: int = 60) -> dict:
        """
        Import a single voucher XML envelope into Tally.
        Returns dict with created/altered/errors counts.
        """
        resp_text = self._send_xml(voucher_xml, timeout=timeout)
        root = self._parse_response(resp_text)

        result = {
            "created": int(root.findtext(".//CREATED", "0")),
            "altered": int(root.findtext(".//ALTERED", "0")),
            "deleted": int(root.findtext(".//DELETED", "0")),
            "errors": int(root.findtext(".//ERRORS", "0")),
            "exceptions": int(root.findtext(".//EXCEPTIONS", "0")),
            "error_message": root.findtext(".//LINEERROR", ""),
        }
        return result

    def import_vouchers_batch(self, voucher_xml_list: list,
                              timeout: int = 120) -> list:
        """
        Import multiple vouchers. Each item is a complete XML envelope string.
        Returns list of result dicts.
        """
        results = []
        for i, xml in enumerate(voucher_xml_list, 1):
            try:
                res = self.import_voucher_xml(xml, timeout=timeout)
                res["index"] = i
                results.append(res)
            except Exception as e:
                results.append({
                    "index": i,
                    "created": 0, "altered": 0, "deleted": 0,
                    "errors": 1, "error_message": str(e),
                })
        return results

    # ── Voucher Export (Read from Tally) ────────────────────────────

    def get_vouchers(self, voucher_type: str, from_date: str, to_date: str) -> list:
        """
        Fetch vouchers of a specific type from Tally for duplicate checking.
        Dates in YYYYMMDD format.

        NOTE: Tally's <CHILDOF> filter is unreliable — it returns ALL voucher types
        regardless of the requested type. VOUCHERTYPE is fetched explicitly and a
        Python-side type filter is applied after parsing.
        """
        xml = f"""<ENVELOPE>
  <HEADER>
    <VERSION>1</VERSION>
    <TALLYREQUEST>Export</TALLYREQUEST>
    <TYPE>Collection</TYPE>
    <ID>Voucher Collection</ID>
  </HEADER>
  <BODY>
    <DESC>
      <STATICVARIABLES>
        <SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
        <SVCURRENTCOMPANY>{self.company}</SVCURRENTCOMPANY>
        <SVFROMDATE>{from_date}</SVFROMDATE>
        <SVTODATE>{to_date}</SVTODATE>
      </STATICVARIABLES>
      <TDL>
        <TDLMESSAGE>
          <COLLECTION NAME="Voucher Collection" ISMODIFY="No">
            <TYPE>Voucher</TYPE>
            <CHILDOF>{voucher_type}</CHILDOF>
            <FETCH>DATE, NARRATION, VOUCHERNUMBER, PARTYLEDGERNAME, AMOUNT, REFERENCE, VOUCHERTYPENAME, PARTYINVNO, PARTYINVDATE</FETCH>
          </COLLECTION>
        </TDLMESSAGE>
      </TDL>
    </DESC>
  </BODY>
</ENVELOPE>"""

        resp_text = self._send_xml(xml)
        root = self._parse_response(resp_text)

        vouchers = []
        target = voucher_type.strip().lower()
        for v in root.iter("VOUCHER"):
            # VOUCHERTYPENAME is the element; VCHTYPE is the attribute — try both
            vtype = (v.findtext("VOUCHERTYPENAME") or v.get("VCHTYPE") or "").strip()
            # Python-side type filter — Tally ignores <CHILDOF> and date filters server-side
            if vtype.lower() != target:
                continue
            vouchers.append({
                "date":         v.findtext("DATE", ""),
                "number":       v.findtext("VOUCHERNUMBER", ""),
                "party":        v.findtext("PARTYLEDGERNAME", ""),
                "amount":       v.findtext("AMOUNT", "0"),
                "narration":    v.findtext("NARRATION", ""),
                "reference":    v.findtext("REFERENCE", ""),
                "vouchertype":  vtype,
                "partyinvno":   v.findtext("PARTYINVNO", ""),
                "partyinvdate": v.findtext("PARTYINVDATE", ""),
            })
        return vouchers

    def get_vouchers_with_guid(self, from_date: str, to_date: str,
                               voucher_type: str = "Purchase") -> list:
        """
        Fetch vouchers with their Tally REMOTEID and VCHKEY for a date range.
        Both are required for targeted deletion via delete_voucher().
        Dates in YYYYMMDD format.
        """
        xml = f"""<ENVELOPE>
  <HEADER>
    <VERSION>1</VERSION>
    <TALLYREQUEST>Export</TALLYREQUEST>
    <TYPE>Collection</TYPE>
    <ID>Voucher GUID Collection</ID>
  </HEADER>
  <BODY>
    <DESC>
      <STATICVARIABLES>
        <SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
        <SVCURRENTCOMPANY>{self.company}</SVCURRENTCOMPANY>
        <SVFROMDATE>{from_date}</SVFROMDATE>
        <SVTODATE>{to_date}</SVTODATE>
      </STATICVARIABLES>
      <TDL>
        <TDLMESSAGE>
          <COLLECTION NAME="Voucher GUID Collection" ISMODIFY="No">
            <TYPE>Voucher</TYPE>
            <CHILDOF>{voucher_type}</CHILDOF>
            <FETCH>MASTERID, GUID, DATE, VOUCHERNUMBER, PARTYLEDGERNAME, AMOUNT</FETCH>
          </COLLECTION>
        </TDLMESSAGE>
      </TDL>
    </DESC>
  </BODY>
</ENVELOPE>"""

        resp_text = self._send_xml(xml)
        root = self._parse_response(resp_text)

        vouchers = []
        for v in root.iter("VOUCHER"):
            # REMOTEID attribute = the voucher's GUID; VCHKEY = compound key needed for delete
            guid = v.get("REMOTEID") or v.findtext("GUID", "") or v.get("GUID", "")
            vchkey = v.get("VCHKEY", "")
            vouchers.append({
                "guid": guid,
                "vchkey": vchkey,
                "date": v.findtext("DATE", ""),
                "number": v.findtext("VOUCHERNUMBER", ""),
                "party": v.findtext("PARTYLEDGERNAME", ""),
                "amount": v.findtext("AMOUNT", "0"),
            })
        return vouchers

    def get_existing_voucher_numbers(self, from_date: str, to_date: str,
                                     voucher_type: str = "Purchase") -> set:
        """
        Return a set of voucher numbers already in Tally for the given range.
        Used as a pre-check before posting to prevent duplicates.
        Dates in YYYYMMDD format.

        NOTE: Tally's TDL SVFROMDATE/SVTODATE filter can be unreliable, so we
        apply a Python-side date filter after fetching to ensure accuracy.
        """
        vouchers = self.get_vouchers_with_guid(from_date, to_date, voucher_type)
        # Python-side date filter — guards against Tally's unreliable TDL date filter
        return {v["number"] for v in vouchers
                if v["number"] and from_date <= v["date"] <= to_date}

    def delete_voucher(self, guid: str, vchkey: str, date: str,
                       voucher_type: str, voucher_number: str,
                       party: str = "") -> dict:
        """
        Delete a single voucher from Tally by its REMOTEID (GUID) + VCHKEY.

        Both guid (REMOTEID) and vchkey are required — Tally Prime needs the
        compound VCHKEY to locate and delete the specific voucher entry.

        Args:
            guid: Tally REMOTEID / GUID of the voucher
            vchkey: Tally VCHKEY (compound key) of the voucher
            date: Voucher date in YYYYMMDD format
            voucher_type: Tally voucher type name (e.g. 'Purchase')
            voucher_number: Voucher number (e.g. 'BILL/2026/03/0001')
            party: Party ledger name (optional, improves matching)

        Returns:
            dict with deleted/errors/error_message keys
        """
        esc_guid = xml_escape(guid)
        esc_vchkey = xml_escape(vchkey)
        esc_type = xml_escape(voucher_type)
        esc_num = xml_escape(voucher_number)
        esc_party = xml_escape(party)
        esc_company = xml_escape(self.company)

        xml = f"""<ENVELOPE>
  <HEADER>
    <VERSION>1</VERSION>
    <TALLYREQUEST>Import</TALLYREQUEST>
    <TYPE>Data</TYPE>
    <ID>Vouchers</ID>
  </HEADER>
  <BODY>
    <DESC>
      <STATICVARIABLES>
        <SVCURRENTCOMPANY>{esc_company}</SVCURRENTCOMPANY>
      </STATICVARIABLES>
    </DESC>
    <DATA>
      <TALLYMESSAGE xmlns:UDF="TallyUDF">
        <VOUCHER VCHTYPE="{esc_type}" ACTION="Delete"
                 REMOTEID="{esc_guid}" VCHKEY="{esc_vchkey}"
                 OBJVIEW="Accounting Voucher View">
          <GUID>{esc_guid}</GUID>
          <DATE>{date}</DATE>
          <VOUCHERNUMBER>{esc_num}</VOUCHERNUMBER>
          <PARTYLEDGERNAME>{esc_party}</PARTYLEDGERNAME>
          <VOUCHERTYPENAME>{esc_type}</VOUCHERTYPENAME>
        </VOUCHER>
      </TALLYMESSAGE>
    </DATA>
  </BODY>
</ENVELOPE>"""

        resp_text = self._send_xml(xml, timeout=30)
        try:
            root = self._parse_response(resp_text)
            return {
                "deleted": int(root.findtext(".//DELETED", "0")),
                "errors": int(root.findtext(".//ERRORS", "0")),
                "error_message": root.findtext(".//LINEERROR", ""),
            }
        except Exception as e:
            return {"deleted": 0, "errors": 1, "error_message": str(e)}

    # ── Day Book (all vouchers for a date range) ────────────────────

    def get_day_book(self, from_date: str, to_date: str) -> str:
        """Export day book XML for a date range (raw XML response)."""
        xml = f"""<ENVELOPE>
  <HEADER>
    <VERSION>1</VERSION>
    <TALLYREQUEST>Export</TALLYREQUEST>
    <TYPE>Data</TYPE>
    <ID>Day Book</ID>
  </HEADER>
  <BODY>
    <DESC>
      <STATICVARIABLES>
        <SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
        <SVCURRENTCOMPANY>{self.company}</SVCURRENTCOMPANY>
        <SVFROMDATE>{from_date}</SVFROMDATE>
        <SVTODATE>{to_date}</SVTODATE>
      </STATICVARIABLES>
    </DESC>
  </BODY>
</ENVELOPE>"""
        return self._send_xml(xml)


# ── CLI Entry Point ─────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description="Tally Prime XML Connector")
    parser.add_argument("--test", action="store_true", help="Test Tally connection")
    parser.add_argument("--list-ledgers", action="store_true",
                        help="List all ledgers in Tally")
    parser.add_argument("--list-groups", action="store_true",
                        help="List all account groups in Tally")
    parser.add_argument("--create-ledger", type=str, default=None,
                        help="Create a ledger (name)")
    parser.add_argument("--group", type=str, default="Sundry Creditors",
                        help="Parent group for new ledger")
    parser.add_argument("--check-ledger", type=str, default=None,
                        help="Check if a ledger exists")

    args = parser.parse_args()
    TMP_DIR.mkdir(parents=True, exist_ok=True)

    tally = TallyConnector()

    if args.test:
        try:
            info = tally.test_connection()
            print(f"✓ Connected to Tally Prime at {info['url']}")
            print(f"  Companies found: {', '.join(info['companies']) if info['companies'] else 'N/A'}")
            print(f"  Target company:  {info['target_company'] or '(not set — set TALLY_COMPANY_NAME in .env)'}")
        except ConnectionError as e:
            print(f"✗ {e}")
            sys.exit(1)

    elif args.list_ledgers:
        ledgers = tally.get_all_ledgers()
        print(f"\n{'Ledger Name':<50} {'Group':<30} {'Opening Bal':>15}")
        print("─" * 95)
        for ldg in ledgers:
            print(f"{ldg['name']:<50} {ldg['group']:<30} {ldg['opening_balance']:>15}")
        print(f"\nTotal: {len(ledgers)} ledgers")
        # Save to tmp
        out = TMP_DIR / "tally_ledgers.json"
        with open(out, "w", encoding="utf-8") as f:
            json.dump(ledgers, f, indent=2)
        print(f"Full list saved to {out}")

    elif args.list_groups:
        groups = tally.get_all_groups()
        print(f"\n{'Group Name':<40} {'Parent':<40}")
        print("─" * 80)
        for g in groups:
            print(f"{g['name']:<40} {g['parent']:<40}")
        print(f"\nTotal: {len(groups)} groups")

    elif args.create_ledger:
        result = tally.create_ledger(args.create_ledger, args.group)
        print(result)

    elif args.check_ledger:
        exists = tally.ledger_exists(args.check_ledger)
        print(f"Ledger '{args.check_ledger}': {'EXISTS' if exists else 'NOT FOUND'}")

    else:
        parser.print_help()


if __name__ == "__main__":
    main()
