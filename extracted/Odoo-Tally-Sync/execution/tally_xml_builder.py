#!/usr/bin/env python3
"""
Tally XML Builder

Transforms Odoo invoice/payment data into Tally Prime-compatible XML
voucher envelopes for import via Tally's XML Server.

Supports:
- Purchase Vouchers (from Odoo vendor bills)
- Sales Vouchers (from Odoo customer invoices)
- Debit Notes (from Odoo vendor returns/refunds)
- Credit Notes (from Odoo customer returns/refunds)
- Payment Vouchers (from Odoo outbound payments)
- Receipt Vouchers (from Odoo inbound payments)

Usage:
    from tally_xml_builder import TallyXMLBuilder
    builder = TallyXMLBuilder(company_name="My Company")
    xml = builder.build_purchase_voucher(odoo_invoice, odoo_lines, mapping)
"""

import json
import re
from datetime import datetime
from pathlib import Path
from typing import Optional
from xml.sax.saxutils import escape as xml_escape


class TallyXMLBuilder:
    """Builds Tally Prime XML voucher envelopes from Odoo data."""

    def __init__(self, company_name: str, vch_prefix: str = ""):
        self.company = company_name
        self.vch_prefix = vch_prefix  # e.g. "TEST " to prefix voucher numbers

    # ── Date Helpers ────────────────────────────────────────────────

    @staticmethod
    def to_tally_date(odoo_date: str) -> str:
        """Convert Odoo date (YYYY-MM-DD) to Tally date (YYYYMMDD)."""
        if not odoo_date:
            return datetime.today().strftime("%Y%m%d")
        return odoo_date.replace("-", "")

    @staticmethod
    def _esc(text: str) -> str:
        """XML-escape text for safe embedding."""
        return xml_escape(str(text)) if text else ""

    # ── Ledger Name Resolution ──────────────────────────────────────

    @staticmethod
    def _strip_account_code(name: str) -> str:
        """
        Strip a leading numeric account code from an Odoo account name.
        e.g. '210501.4 Staff Welfare' -> 'Staff Welfare'
             '112320 SGST Payable'    -> 'SGST Payable'
        Returns the original name if no numeric prefix is found.
        """
        stripped = re.sub(r'^\d[\d.]*\s+', '', name)
        return stripped if stripped else name

    @staticmethod
    def resolve_ledger(name: str, mapping: dict) -> str:
        """
        Look up the Tally ledger name from the mapping.
        If not found, strips any leading numeric account code so that
        Odoo accounts like '210501.4 Staff Welfare' resolve to the
        existing Tally ledger 'Staff Welfare' instead of creating a
        duplicate with the code prefix.
        """
        ledger_map = mapping.get("odoo_to_tally_ledgers", {})
        if name in ledger_map:
            return ledger_map[name]
        stripped = re.sub(r'^\d[\d.]*\s+', '', name)
        if stripped and stripped != name:
            return ledger_map.get(stripped, stripped)
        return name

    @staticmethod
    def resolve_partner(name: str, mapping: dict) -> str:
        """Look up the Tally party ledger name from the mapping."""
        partner_map = mapping.get("odoo_to_tally_partners", {})
        return partner_map.get(name, name)

    # ── Voucher Envelope Wrapper ────────────────────────────────────

    def _wrap_envelope(self, voucher_xml: str) -> str:
        """Wrap a <VOUCHER> block in the Tally import envelope."""
        return f"""<ENVELOPE>
  <HEADER>
    <VERSION>1</VERSION>
    <TALLYREQUEST>Import</TALLYREQUEST>
    <TYPE>Data</TYPE>
    <ID>Vouchers</ID>
  </HEADER>
  <BODY>
    <DESC>
      <STATICVARIABLES>
        <SVCURRENTCOMPANY>{self._esc(self.company)}</SVCURRENTCOMPANY>
      </STATICVARIABLES>
    </DESC>
    <DATA>
      <TALLYMESSAGE>
{voucher_xml}
      </TALLYMESSAGE>
    </DATA>
  </BODY>
</ENVELOPE>"""

    # ── Purchase Voucher ────────────────────────────────────────────

    # ── Generic Ledger Entries Builder ─────────────────────────────

    # Odoo accounts that need product-based routing to Tally ledgers
    DYNAMIC_SALES_ACCOUNTS = {"200110 Local Sales"}

    # Tally ledger names that represent discount-allowed (absorbed into net sales)
    DISCOUNT_ABSORB_NAMES = {"discount allowed"}

    # Odoo GST Payable accounts that need rate-aware routing for sales
    # Includes both "CODE Name" and plain "Name" formats since Odoo
    # credit notes sometimes return account names without the numeric prefix.
    OUTPUT_GST_MAP = {
        "112320 SGST Payable": {9: "Output SGST @ 9%", 6: "Output SGST @ 6%", 2.5: "Output SGST @ 2.5%"},
        "SGST Payable":        {9: "Output SGST @ 9%", 6: "Output SGST @ 6%", 2.5: "Output SGST @ 2.5%"},
        "112330 CGST Payable": {9: "Output CGST @ 9%", 6: "Output CGST @ 6%", 2.5: "Output CGST @ 2.5%"},
        "CGST Payable":        {9: "Output CGST @ 9%", 6: "Output CGST @ 6%", 2.5: "Output CGST @ 2.5%"},
        "112340 IGST Payable": {18: "Output IGST @ 18%", 12: "Output IGST @ 12%", 5: "Output IGST @ 5%"},
        "IGST Payable":        {18: "Output IGST @ 18%", 12: "Output IGST @ 12%", 5: "Output IGST @ 5%"},
    }

    @staticmethod
    def _resolve_output_gst_ledger(line: dict, account_name: str, rate_map: dict) -> str:
        """
        Resolve the correct Output GST Tally ledger based on the tax rate
        embedded in the tax_line_id name (e.g. 'SGST Sale 9%' → 9).
        """
        tax_line = line.get("tax_line_id")
        if isinstance(tax_line, (list, tuple)):
            tax_name = tax_line[1]
        else:
            tax_name = str(tax_line or "")

        # Extract percentage from tax name like "SGST Sale 9%" or "IGST 18%"
        import re
        match = re.search(r'(\d+(?:\.\d+)?)\s*%', tax_name)
        if match:
            rate = float(match.group(1))
            if rate in rate_map:
                return rate_map[rate]

        # Fallback: try the most common rate (9% for SGST/CGST, 18% for IGST)
        if "IGST" in account_name:
            return rate_map.get(18, list(rate_map.values())[0])
        return rate_map.get(9, list(rate_map.values())[0])

    @staticmethod
    def _resolve_dynamic_sales_ledger(line: dict) -> str:
        """
        Route '200110 Local Sales' to the correct Tally ledger based on
        the product info attached to the line:
          - Product name contains 'PrintStick' (case-insensitive) → Printsticks
          - Product type is 'service' → Sale of Services
          - Everything else (consu/product) → Sale of Goods
        """
        product = line.get("product_id")
        if not product:
            return "Sale of Goods"  # fallback for lines without a product

        prod_name = product[1] if isinstance(product, (list, tuple)) else str(product)
        if "printstick" in prod_name.lower():
            return "Printsticks"

        prod_type = line.get("_product_type", "")  # enriched by sync pipeline
        if prod_type == "service":
            return "Sale of Services"

        return "Sale of Goods"

    @staticmethod
    def _is_payable_or_receivable(line: dict) -> bool:
        """Check if a journal line is the payable/receivable (party) line."""
        acct_type = line.get("account_type", "")
        return acct_type in ("liability_payable", "asset_receivable")

    def _build_ledger_entries(self, lines: list, mapping: dict,
                               party_ledger: str,
                               is_sales: bool = False) -> str:
        """
        Build ALLLEDGERENTRIES.LIST XML from Odoo journal entry lines.

        Processes ALL lines — payable/receivable lines use the party ledger
        name, all others use the resolved account ledger name. Handles both
        debit and credit sides so TDS, rounding, and other adjustments are
        never dropped.

        Lines with the same ledger name are consolidated into a single entry
        by netting debit and credit amounts to avoid duplicate line items.

        When is_sales=True, GST Payable accounts are routed to Output GST
        ledgers (rate-specific) instead of Input GST ledgers.
        """
        # Consolidate: {ledger_name: net_balance}
        # Positive = credit (Cr), Negative = debit (Dr)
        consolidated = {}
        for line in lines:
          debit = line.get("debit", 0)
          credit = line.get("credit", 0)
          if debit == 0 and credit == 0:
            continue

          if self._is_payable_or_receivable(line):
            ledger_name = party_ledger
          else:
            account_name = (line["account_id"][1]
                    if isinstance(line.get("account_id"), (list, tuple))
                    else str(line.get("account_id", "")))
            # Dynamic routing for revenue accounts like '200110 Local Sales'
            if account_name in self.DYNAMIC_SALES_ACCOUNTS:
              ledger_name = self._resolve_dynamic_sales_ledger(line)
            # Output GST routing for sales-type vouchers
            elif is_sales and account_name in self.OUTPUT_GST_MAP:
              ledger_name = self._resolve_output_gst_ledger(
                line, account_name, self.OUTPUT_GST_MAP[account_name])
            else:
              ledger_name = self.resolve_ledger(account_name, mapping)

          # Net balance: credit positive, debit negative
          consolidated[ledger_name] = consolidated.get(ledger_name, 0) + credit - debit

        # For sales: absorb Discount Allowed into the primary sales ledger
        # so that Tally shows net sales amount without a separate discount line
        if is_sales:
          # Match discount ledgers by substring so "400100 Discount Allowed"
          # is caught even if the code prefix was not stripped by the mapping.
          discount_keys = [k for k in list(consolidated)
                           if any(d in k.lower() for d in self.DISCOUNT_ABSORB_NAMES)]
          if discount_keys:
            # Find first positive (credit) non-party, non-tax revenue entry
            _skip = {'cgst', 'sgst', 'igst', 'gst', 'tds', 'tax', 'round', 'cess', 'discount'}
            revenue_key = next(
              (k for k, v in consolidated.items()
               if v > 0 and k != party_ledger
               and not any(t in k.lower() for t in _skip)),
              None
            )
            if revenue_key:
              for dk in discount_keys:
                consolidated[revenue_key] += consolidated.pop(dk)

        # Build XML from consolidated entries
        # Party ledger always first (debit for sales, credit for purchase)
        entries = ""
        if party_ledger in consolidated:
          net = consolidated[party_ledger]
          if abs(net) >= 0.001:
            is_deemed = "Yes" if net < 0 else "No"
            entries += f"""
        <ALLLEDGERENTRIES.LIST>
          <LEDGERNAME>{self._esc(party_ledger)}</LEDGERNAME>
          <ISDEEMEDPOSITIVE>{is_deemed}</ISDEEMEDPOSITIVE>
          <AMOUNT>{round(net, 2)}</AMOUNT>
        </ALLLEDGERENTRIES.LIST>"""
        # Other ledgers
        for ledger_name, net in consolidated.items():
          if ledger_name == party_ledger:
            continue
          if abs(net) < 0.001:
            continue  # Skip zero-net entries
          if net >= 0:
            # Net credit
            entries += f"""
        <ALLLEDGERENTRIES.LIST>
          <LEDGERNAME>{self._esc(ledger_name)}</LEDGERNAME>
          <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>
          <AMOUNT>{round(net, 2)}</AMOUNT>
        </ALLLEDGERENTRIES.LIST>"""
          else:
            # Net debit for non-party ledgers (rare, but preserve)
            entries += f"""
        <ALLLEDGERENTRIES.LIST>
          <LEDGERNAME>{self._esc(ledger_name)}</LEDGERNAME>
          <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>
          <AMOUNT>{round(net, 2)}</AMOUNT>
        </ALLLEDGERENTRIES.LIST>"""
        return entries

    # ── Purchase Voucher ────────────────────────────────────────────

    def build_purchase_voucher(self, invoice: dict, lines: list,
                                mapping: dict) -> str:
        """
        Build a Tally Purchase voucher XML from an Odoo vendor bill.

        Args:
            invoice: Odoo account.move record (in_invoice)
            lines: Odoo account.move.line records for this invoice
            mapping: Ledger mapping dict
        """
        partner_name = invoice["partner_id"][1] if invoice.get("partner_id") else "Unknown"
        tally_party = self.resolve_partner(partner_name, mapping)
        # Voucher Date = Odoo Accounting Date (date); Supplier Invoice Date = Odoo Bill Date (invoice_date)
        voucher_date = self.to_tally_date(invoice.get("date", "") or invoice.get("invoice_date", ""))
        supplier_inv_date = self.to_tally_date(invoice.get("invoice_date", ""))
        bill_number = invoice.get("name", "")
        vendor_ref = invoice.get("ref", "") or ""
        narration = f"Odoo Bill: {bill_number}"
        if vendor_ref:
            narration += f" | Vendor Invoice: {vendor_ref}"

        ledger_entries = self._build_ledger_entries(lines, mapping, tally_party)

        voucher = f"""
        <VOUCHER VCHTYPE="Purchase" ACTION="Create">
          <DATE>{voucher_date}</DATE>
          <VOUCHERTYPENAME>Purchase</VOUCHERTYPENAME>
          <PERSISTEDVIEW>Invoice Voucher View</PERSISTEDVIEW>
          <ISINVOICE>Yes</ISINVOICE>
          <PARTYLEDGERNAME>{self._esc(tally_party)}</PARTYLEDGERNAME>
          <NARRATION>{self._esc(narration)}</NARRATION>
          <REFERENCE>{self._esc(vendor_ref if vendor_ref else bill_number)}</REFERENCE>
          <VOUCHERNUMBER>{self._esc(self.vch_prefix + bill_number)}</VOUCHERNUMBER>
          <PARTYINVNO>{self._esc(vendor_ref if vendor_ref else bill_number)}</PARTYINVNO>
          <PARTYINVDATE>{supplier_inv_date}</PARTYINVDATE>
          {ledger_entries}
        </VOUCHER>"""

        return self._wrap_envelope(voucher)

    # ── Sales Voucher ───────────────────────────────────────────────

    def build_sales_voucher(self, invoice: dict, lines: list,
                             mapping: dict) -> str:
        """Build a Tally Sales voucher XML from an Odoo customer invoice."""
        partner_name = invoice["partner_id"][1] if invoice.get("partner_id") else "Unknown"
        tally_party = self.resolve_partner(partner_name, mapping)
        inv_date = self.to_tally_date(invoice.get("invoice_date", ""))
        bill_number = invoice.get("name", "")
        vendor_ref = invoice.get("ref", "") or ""
        narration = f"Odoo Ref: {bill_number}"
        if vendor_ref:
            narration += f" | Customer Ref: {vendor_ref}"

        ledger_entries = self._build_ledger_entries(lines, mapping, tally_party, is_sales=True)

        voucher = f"""
        <VOUCHER VCHTYPE="Sales" ACTION="Create">
          <DATE>{inv_date}</DATE>
          <VOUCHERTYPENAME>Sales</VOUCHERTYPENAME>
          <PERSISTEDVIEW>Accounting Voucher View</PERSISTEDVIEW>
          <PARTYLEDGERNAME>{self._esc(tally_party)}</PARTYLEDGERNAME>
          <NARRATION>{self._esc(narration)}</NARRATION>
          <REFERENCE>{self._esc(vendor_ref if vendor_ref else bill_number)}</REFERENCE>
          <VOUCHERNUMBER>{self._esc(self.vch_prefix + bill_number)}</VOUCHERNUMBER>
          {ledger_entries}
        </VOUCHER>"""

        return self._wrap_envelope(voucher)

    # ── GST Invoice Voucher ───────────────────────────────────────

    def build_gst_invoice_voucher(self, invoice: dict, lines: list,
                                   mapping: dict) -> str:
        """Build a Tally GST INVOICE voucher XML from an Odoo customer invoice."""
        partner_name = invoice["partner_id"][1] if invoice.get("partner_id") else "Unknown"
        tally_party = self.resolve_partner(partner_name, mapping)
        inv_date = self.to_tally_date(invoice.get("invoice_date", ""))
        bill_number = invoice.get("name", "")
        vendor_ref = invoice.get("ref", "") or ""
        narration = f"Odoo Ref: {bill_number}"
        if vendor_ref:
            narration += f" | Customer Ref: {vendor_ref}"

        ledger_entries = self._build_ledger_entries(lines, mapping, tally_party, is_sales=True)

        voucher = f"""
        <VOUCHER VCHTYPE="GST INVOICE" ACTION="Create">
          <DATE>{inv_date}</DATE>
          <VOUCHERTYPENAME>GST INVOICE</VOUCHERTYPENAME>
          <PERSISTEDVIEW>Accounting Voucher View</PERSISTEDVIEW>
          <PARTYLEDGERNAME>{self._esc(tally_party)}</PARTYLEDGERNAME>
          <NARRATION>{self._esc(narration)}</NARRATION>
          <REFERENCE>{self._esc(vendor_ref if vendor_ref else bill_number)}</REFERENCE>
          <VOUCHERNUMBER>{self._esc(self.vch_prefix + bill_number)}</VOUCHERNUMBER>
          {ledger_entries}
        </VOUCHER>"""

        return self._wrap_envelope(voucher)

    # ── Debit Note (Purchase Return) ────────────────────────────────

    def build_debit_note(self, invoice: dict, lines: list,
                          mapping: dict) -> str:
        """Build a Tally Debit Note voucher from an Odoo vendor refund."""
        partner_name = invoice["partner_id"][1] if invoice.get("partner_id") else "Unknown"
        tally_party = self.resolve_partner(partner_name, mapping)
        voucher_date = self.to_tally_date(invoice.get("date", "") or invoice.get("invoice_date", ""))
        supplier_inv_date = self.to_tally_date(invoice.get("invoice_date", ""))
        bill_number = invoice.get("name", "")
        vendor_ref = invoice.get("ref", "") or ""
        narration = f"Odoo Debit Note: {bill_number}"
        if vendor_ref:
            narration += f" | Vendor Ref: {vendor_ref}"

        ledger_entries = self._build_ledger_entries(lines, mapping, tally_party)

        voucher = f"""
        <VOUCHER VCHTYPE="Debit Note" ACTION="Create">
          <DATE>{voucher_date}</DATE>
          <VOUCHERTYPENAME>Debit Note</VOUCHERTYPENAME>
          <PERSISTEDVIEW>Invoice Voucher View</PERSISTEDVIEW>
          <ISINVOICE>Yes</ISINVOICE>
          <PARTYLEDGERNAME>{self._esc(tally_party)}</PARTYLEDGERNAME>
          <NARRATION>{self._esc(narration)}</NARRATION>
          <REFERENCE>{self._esc(vendor_ref if vendor_ref else bill_number)}</REFERENCE>
          <VOUCHERNUMBER>{self._esc(self.vch_prefix + bill_number)}</VOUCHERNUMBER>
          <PARTYINVNO>{self._esc(vendor_ref if vendor_ref else bill_number)}</PARTYINVNO>
          <PARTYINVDATE>{supplier_inv_date}</PARTYINVDATE>
          {ledger_entries}
        </VOUCHER>"""

        return self._wrap_envelope(voucher)

    # ── Credit Note (Sales Return) ──────────────────────────────────

    def build_credit_note(self, invoice: dict, lines: list,
                           mapping: dict) -> str:
        """Build a Tally Credit Note voucher from an Odoo customer refund."""
        partner_name = invoice["partner_id"][1] if invoice.get("partner_id") else "Unknown"
        tally_party = self.resolve_partner(partner_name, mapping)
        inv_date = self.to_tally_date(invoice.get("invoice_date", ""))
        bill_number = invoice.get("name", "")
        vendor_ref = invoice.get("ref", "") or ""
        narration = f"Odoo Credit Note: {bill_number}"
        if vendor_ref:
            narration += f" | Customer Ref: {vendor_ref}"

        ledger_entries = self._build_ledger_entries(lines, mapping, tally_party, is_sales=True)

        voucher = f"""
        <VOUCHER VCHTYPE="Credit Note" ACTION="Create">
          <DATE>{inv_date}</DATE>
          <VOUCHERTYPENAME>Credit Note</VOUCHERTYPENAME>
          <PERSISTEDVIEW>Accounting Voucher View</PERSISTEDVIEW>
          <PARTYLEDGERNAME>{self._esc(tally_party)}</PARTYLEDGERNAME>
          <NARRATION>{self._esc(narration)}</NARRATION>
          <REFERENCE>{self._esc(vendor_ref if vendor_ref else bill_number)}</REFERENCE>
          <VOUCHERNUMBER>{self._esc(self.vch_prefix + bill_number)}</VOUCHERNUMBER>
          {ledger_entries}
        </VOUCHER>"""

        return self._wrap_envelope(voucher)

    # ── Payment Voucher (Odoo → Tally) ──────────────────────────────

    def build_payment_voucher(self, payment: dict, mapping: dict,
                               bank_ledger: str = "Bank Accounts") -> str:
        """Build a Tally Payment voucher from an Odoo outbound payment."""
        partner_name = payment["partner_id"][1] if payment.get("partner_id") else "Unknown"
        tally_party = self.resolve_partner(partner_name, mapping)
        pay_date = self.to_tally_date(payment.get("date", ""))
        ref = payment.get("name", "")
        amount = payment.get("amount", 0)
        narration = f"Odoo Payment Ref: {ref}"

        tally_bank = self.resolve_ledger(bank_ledger, mapping)

        voucher = f"""
        <VOUCHER VCHTYPE="Payment" ACTION="Create">
          <DATE>{pay_date}</DATE>
          <VOUCHERTYPENAME>Payment</VOUCHERTYPENAME>
          <PERSISTEDVIEW>Accounting Voucher View</PERSISTEDVIEW>
          <PARTYLEDGERNAME>{self._esc(tally_party)}</PARTYLEDGERNAME>
          <NARRATION>{self._esc(narration)}</NARRATION>
          <REFERENCE>{self._esc(ref)}</REFERENCE>
          <VOUCHERNUMBER>{self._esc(self.vch_prefix + ref)}</VOUCHERNUMBER>
          <ALLLEDGERENTRIES.LIST>
            <LEDGERNAME>{self._esc(tally_party)}</LEDGERNAME>
            <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>
            <AMOUNT>-{amount}</AMOUNT>
          </ALLLEDGERENTRIES.LIST>
          <ALLLEDGERENTRIES.LIST>
            <LEDGERNAME>{self._esc(tally_bank)}</LEDGERNAME>
            <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>
            <AMOUNT>{amount}</AMOUNT>
          </ALLLEDGERENTRIES.LIST>
        </VOUCHER>"""

        return self._wrap_envelope(voucher)

    # ── Receipt Voucher (Odoo → Tally) ──────────────────────────────

    def build_receipt_voucher(self, payment: dict, mapping: dict,
                               bank_ledger: str = "Bank Accounts") -> str:
        """Build a Tally Receipt voucher from an Odoo inbound payment."""
        partner_name = payment["partner_id"][1] if payment.get("partner_id") else "Unknown"
        tally_party = self.resolve_partner(partner_name, mapping)
        pay_date = self.to_tally_date(payment.get("date", ""))
        ref = payment.get("name", "")
        amount = payment.get("amount", 0)
        narration = f"Odoo Receipt Ref: {ref}"

        tally_bank = self.resolve_ledger(bank_ledger, mapping)

        voucher = f"""
        <VOUCHER VCHTYPE="Receipt" ACTION="Create">
          <DATE>{pay_date}</DATE>
          <VOUCHERTYPENAME>Receipt</VOUCHERTYPENAME>
          <PERSISTEDVIEW>Accounting Voucher View</PERSISTEDVIEW>
          <PARTYLEDGERNAME>{self._esc(tally_party)}</PARTYLEDGERNAME>
          <NARRATION>{self._esc(narration)}</NARRATION>
          <REFERENCE>{self._esc(ref)}</REFERENCE>
          <VOUCHERNUMBER>{self._esc(self.vch_prefix + ref)}</VOUCHERNUMBER>
          <ALLLEDGERENTRIES.LIST>
            <LEDGERNAME>{self._esc(tally_bank)}</LEDGERNAME>
            <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>
            <AMOUNT>-{amount}</AMOUNT>
          </ALLLEDGERENTRIES.LIST>
          <ALLLEDGERENTRIES.LIST>
            <LEDGERNAME>{self._esc(tally_party)}</LEDGERNAME>
            <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>
            <AMOUNT>{amount}</AMOUNT>
          </ALLLEDGERENTRIES.LIST>
        </VOUCHER>"""

        return self._wrap_envelope(voucher)
