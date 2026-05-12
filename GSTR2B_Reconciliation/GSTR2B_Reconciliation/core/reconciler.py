"""
GSTR2B Reconciliation Engine.

Performs multi-level matching between Tally and GSTR2B data
and classifies records into reconciliation buckets.
"""
import logging
from dataclasses import dataclass, field
from datetime import datetime

import pandas as pd
import numpy as np

from core.utils import (
    normalize_invoice_number, fuzzy_name_score, amounts_match,
    dates_within_tolerance, validate_gstin, safe_float,
)
import config

logger = logging.getLogger("gstr2b_reco")


# ---------------------------------------------------------------------------
# Result Container
# ---------------------------------------------------------------------------
@dataclass
class ReconciliationResult:
    """Holds all reconciliation outputs."""
    matched: pd.DataFrame = field(default_factory=pd.DataFrame)
    in_tally_not_gstr2b: pd.DataFrame = field(default_factory=pd.DataFrame)
    in_gstr2b_not_tally: pd.DataFrame = field(default_factory=pd.DataFrame)
    no_bill_no_payment: pd.DataFrame = field(default_factory=pd.DataFrame)
    amount_mismatch: pd.DataFrame = field(default_factory=pd.DataFrame)
    duplicates_tally: pd.DataFrame = field(default_factory=pd.DataFrame)
    duplicates_gstr2b: pd.DataFrame = field(default_factory=pd.DataFrame)
    errors: list = field(default_factory=list)
    summary: dict = field(default_factory=dict)
    run_time_seconds: float = 0.0


# ---------------------------------------------------------------------------
# Reconciler
# ---------------------------------------------------------------------------
class Reconciler:
    """
    Main reconciliation class.

    Usage:
        reconciler = Reconciler(tally_df, gstr2b_df)
        result = reconciler.run()
    """

    def __init__(
        self,
        tally_df: pd.DataFrame,
        gstr2b_df: pd.DataFrame,
        fuzzy_threshold: int = config.FUZZY_MATCH_THRESHOLD,
        date_tolerance: int = config.DATE_TOLERANCE_DAYS,
        amt_tolerance_pct: float = config.AMOUNT_TOLERANCE_PERCENT,
        amt_tolerance_abs: float = config.AMOUNT_TOLERANCE_ABSOLUTE,
    ):
        self.tally = tally_df.copy()
        self.gstr2b = gstr2b_df.copy()
        self.fuzzy_threshold = fuzzy_threshold
        self.date_tolerance = date_tolerance
        self.amt_tol_pct = amt_tolerance_pct
        self.amt_tol_abs = amt_tolerance_abs

        # Tracking indices
        self.tally["_idx"] = range(len(self.tally))
        self.gstr2b["_idx"] = range(len(self.gstr2b))
        self._matched_tally_idx = set()
        self._matched_gstr2b_idx = set()
        self._matches = []       # list of (tally_idx, gstr2b_idx, match_type, score)
        self._amount_mismatches = []
        self._errors = []

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------
    def run(self) -> ReconciliationResult:
        """Execute full reconciliation pipeline."""
        start = datetime.now()
        logger.info("=" * 60)
        logger.info("Starting GSTR2B Reconciliation")
        logger.info(f"Tally records : {len(self.tally)}")
        logger.info(f"GSTR2B records: {len(self.gstr2b)}")
        logger.info("=" * 60)

        # Step 1: Detect duplicates
        dup_tally = self._find_duplicates(self.tally, "tally")
        dup_gstr2b = self._find_duplicates(self.gstr2b, "gstr2b")

        # Step 2: Primary matching — GSTIN + Invoice Number (exact)
        self._match_primary()

        # Step 3: Secondary matching — fuzzy on unmatched
        self._match_secondary()

        # Step 4: Build result DataFrames
        result = self._build_results(dup_tally, dup_gstr2b)

        elapsed = (datetime.now() - start).total_seconds()
        result.run_time_seconds = elapsed

        # Summary statistics
        result.summary = {
            "tally_total_records": len(self.tally),
            "gstr2b_total_records": len(self.gstr2b),
            "matched_count": len(result.matched),
            "in_tally_not_gstr2b": len(result.in_tally_not_gstr2b),
            "in_gstr2b_not_tally": len(result.in_gstr2b_not_tally),
            "no_bill_no_payment": len(result.no_bill_no_payment),
            "amount_mismatch": len(result.amount_mismatch),
            "duplicates_tally": len(result.duplicates_tally),
            "duplicates_gstr2b": len(result.duplicates_gstr2b),
            "errors_count": len(result.errors),
            "run_time_seconds": round(elapsed, 2),
            "tally_itc_total": round(self.tally["total_tax"].sum(), 2),
            "gstr2b_itc_total": round(self.gstr2b["total_tax"].sum(), 2),
        }

        logger.info(f"Reconciliation complete in {elapsed:.2f}s")
        for k, v in result.summary.items():
            logger.info(f"  {k}: {v}")

        return result

    # ------------------------------------------------------------------
    # Primary Match: GSTIN + Normalized Invoice Number
    # ------------------------------------------------------------------
    def _match_primary(self):
        """Exact match on GSTIN + normalized invoice number."""
        logger.info("Phase 1: Primary matching (GSTIN + Invoice Number)")

        # Build lookup from GSTR2B
        gstr2b_lookup = {}
        for _, row in self.gstr2b.iterrows():
            if row["gstin"] and row["invoice_number"]:
                key = (row["gstin"], row["invoice_number"])
                gstr2b_lookup.setdefault(key, []).append(row["_idx"])

        count = 0
        for _, trow in self.tally.iterrows():
            if trow["_idx"] in self._matched_tally_idx:
                continue
            if not trow["gstin"] or not trow["bill_number"]:
                continue

            key = (trow["gstin"], trow["bill_number"])
            if key in gstr2b_lookup:
                for gidx in gstr2b_lookup[key]:
                    if gidx in self._matched_gstr2b_idx:
                        continue
                    grow = self.gstr2b[self.gstr2b["_idx"] == gidx].iloc[0]

                    # Check if amounts match
                    tax_match = amounts_match(
                        trow["total_tax"], grow["total_tax"],
                        self.amt_tol_pct, self.amt_tol_abs
                    )
                    val_match = amounts_match(
                        trow["taxable_value"], grow["taxable_value"],
                        self.amt_tol_pct, self.amt_tol_abs
                    )

                    if tax_match and val_match:
                        self._matches.append((trow["_idx"], gidx, "EXACT", 100))
                    else:
                        self._matches.append((trow["_idx"], gidx, "EXACT_AMT_MISMATCH", 90))
                        self._amount_mismatches.append((trow["_idx"], gidx))

                    self._matched_tally_idx.add(trow["_idx"])
                    self._matched_gstr2b_idx.add(gidx)
                    count += 1
                    break

        logger.info(f"  Primary matches: {count}")

    # ------------------------------------------------------------------
    # Secondary Match: Fuzzy Name + Date + Amount
    # ------------------------------------------------------------------
    def _match_secondary(self):
        """Fuzzy matching on unmatched records."""
        logger.info("Phase 2: Secondary matching (Fuzzy name + Date + Amount)")

        unmatched_tally = self.tally[~self.tally["_idx"].isin(self._matched_tally_idx)]
        unmatched_gstr2b = self.gstr2b[~self.gstr2b["_idx"].isin(self._matched_gstr2b_idx)]

        if unmatched_tally.empty or unmatched_gstr2b.empty:
            logger.info("  No unmatched records to process")
            return

        count = 0

        # Group GSTR2B by GSTIN for faster lookup
        gstr2b_by_gstin = {}
        for _, grow in unmatched_gstr2b.iterrows():
            gstr2b_by_gstin.setdefault(grow["gstin"], []).append(grow)

        for _, trow in unmatched_tally.iterrows():
            if trow["_idx"] in self._matched_tally_idx:
                continue

            best_match = None
            best_score = 0

            # Try GSTIN-based candidates first
            candidates = gstr2b_by_gstin.get(trow["gstin"], []) if trow["gstin"] else []

            # Also try fuzzy name matching across all unmatched
            if not candidates:
                for _, grow in unmatched_gstr2b.iterrows():
                    if grow["_idx"] in self._matched_gstr2b_idx:
                        continue
                    name_score = fuzzy_name_score(trow["party_name"], grow["supplier_name"])
                    if name_score >= self.fuzzy_threshold:
                        candidates.append(grow)

            for grow in candidates:
                if grow["_idx"] in self._matched_gstr2b_idx:
                    continue

                score = 0

                # GSTIN match
                if trow["gstin"] and trow["gstin"] == grow["gstin"]:
                    score += 30

                # Name similarity
                name_score = fuzzy_name_score(trow["party_name"], grow["supplier_name"])
                if name_score >= self.fuzzy_threshold:
                    score += (name_score / 100) * 20

                # Invoice number similarity
                if trow["bill_number"] and grow["invoice_number"]:
                    if trow["bill_number"] == grow["invoice_number"]:
                        score += 20
                    elif trow["bill_number"] in grow["invoice_number"] or grow["invoice_number"] in trow["bill_number"]:
                        score += 10

                # Date proximity
                if dates_within_tolerance(trow.get("voucher_date"), grow.get("invoice_date"), self.date_tolerance):
                    score += 15

                # Amount match
                if amounts_match(trow["taxable_value"], grow["taxable_value"], self.amt_tol_pct, self.amt_tol_abs):
                    score += 15

                if score > best_score and score >= 50:
                    best_score = score
                    best_match = grow

            if best_match is not None:
                gidx = best_match["_idx"]

                # Check amount mismatch
                tax_match = amounts_match(
                    trow["total_tax"], best_match["total_tax"],
                    self.amt_tol_pct, self.amt_tol_abs
                )
                match_type = "FUZZY" if tax_match else "FUZZY_AMT_MISMATCH"
                if not tax_match:
                    self._amount_mismatches.append((trow["_idx"], gidx))

                self._matches.append((trow["_idx"], gidx, match_type, best_score))
                self._matched_tally_idx.add(trow["_idx"])
                self._matched_gstr2b_idx.add(gidx)
                count += 1

        logger.info(f"  Secondary matches: {count}")

    # ------------------------------------------------------------------
    # Duplicate Detection
    # ------------------------------------------------------------------
    def _find_duplicates(self, df: pd.DataFrame, source: str) -> pd.DataFrame:
        """Find duplicate entries based on GSTIN + Invoice Number."""
        if source == "tally":
            key_cols = ["gstin", "bill_number"]
        else:
            key_cols = ["gstin", "invoice_number"]

        # Only check rows that have both key fields
        mask = (df[key_cols[0]] != "") & (df[key_cols[1]] != "")
        subset = df[mask]

        if subset.empty:
            return pd.DataFrame()

        dupes = subset[subset.duplicated(subset=key_cols, keep=False)].copy()
        if not dupes.empty:
            logger.warning(f"Found {len(dupes)} potential duplicate entries in {source}")
            dupes["duplicate_flag"] = "DUPLICATE"

        return dupes

    # ------------------------------------------------------------------
    # Build Result DataFrames
    # ------------------------------------------------------------------
    def _build_results(self, dup_tally, dup_gstr2b) -> ReconciliationResult:
        """Classify all records and build output DataFrames."""
        result = ReconciliationResult()
        result.errors = self._errors

        # ---- Matched ----
        if self._matches:
            matched_rows = []
            for tidx, gidx, match_type, score in self._matches:
                trow = self.tally[self.tally["_idx"] == tidx].iloc[0]
                grow = self.gstr2b[self.gstr2b["_idx"] == gidx].iloc[0]
                matched_rows.append({
                    "Vendor Name (Tally)": trow["party_name"],
                    "Supplier Name (GSTR2B)": grow["supplier_name"],
                    "GSTIN": trow["gstin"] or grow["gstin"],
                    "Invoice No (Tally)": trow.get("bill_number_raw", trow["bill_number"]),
                    "Invoice No (GSTR2B)": grow.get("invoice_number_raw", grow["invoice_number"]),
                    "Invoice Date (Tally)": trow.get("voucher_date", ""),
                    "Invoice Date (GSTR2B)": grow.get("invoice_date", ""),
                    "Taxable Value (Tally)": trow["taxable_value"],
                    "Taxable Value (GSTR2B)": grow["taxable_value"],
                    "IGST (Tally)": trow["igst"],
                    "IGST (GSTR2B)": grow["igst"],
                    "CGST (Tally)": trow["cgst"],
                    "CGST (GSTR2B)": grow["cgst"],
                    "SGST (Tally)": trow["sgst"],
                    "SGST (GSTR2B)": grow["sgst"],
                    "Total Tax (Tally)": trow["total_tax"],
                    "Total Tax (GSTR2B)": grow["total_tax"],
                    "Tax Difference": round(trow["total_tax"] - grow["total_tax"], 2),
                    "Match Type": match_type,
                    "Match Score": score,
                    "Status": "Matched" if "MISMATCH" not in match_type else "Amount Mismatch",
                })
            result.matched = pd.DataFrame(matched_rows)

        # ---- Amount Mismatches (subset of matched) ----
        if result.matched is not None and not result.matched.empty:
            result.amount_mismatch = result.matched[
                result.matched["Status"] == "Amount Mismatch"
            ].copy()

        # ---- In Tally but NOT in GSTR2B ----
        unmatched_tally = self.tally[~self.tally["_idx"].isin(self._matched_tally_idx)]
        if not unmatched_tally.empty:
            rows = []
            for _, r in unmatched_tally.iterrows():
                rows.append({
                    "Vendor Name": r["party_name"],
                    "GSTIN": r["gstin"],
                    "Invoice Number": r.get("bill_number_raw", r["bill_number"]),
                    "Invoice Date": r.get("voucher_date", ""),
                    "Voucher Type": r.get("voucher_type", ""),
                    "Voucher Number": r.get("voucher_number", ""),
                    "Taxable Value": r["taxable_value"],
                    "IGST": r["igst"],
                    "CGST": r["cgst"],
                    "SGST": r["sgst"],
                    "Total Tax": r["total_tax"],
                    "Total Value": r["total_value"],
                    "Narration": r.get("narration", ""),
                    "Source": "Tally",
                    "GSTIN Valid": "Yes" if validate_gstin(r["gstin"]) else "No",
                    "Status": "In Tally, Missing in GSTR2B",
                })
            result.in_tally_not_gstr2b = pd.DataFrame(rows)

        # ---- In GSTR2B but NOT in Tally ----
        unmatched_gstr2b = self.gstr2b[~self.gstr2b["_idx"].isin(self._matched_gstr2b_idx)]
        if not unmatched_gstr2b.empty:
            rows = []
            for _, r in unmatched_gstr2b.iterrows():
                rows.append({
                    "Supplier Name": r["supplier_name"],
                    "GSTIN": r["gstin"],
                    "Invoice Number": r.get("invoice_number_raw", r["invoice_number"]),
                    "Invoice Date": r.get("invoice_date", ""),
                    "Taxable Value": r["taxable_value"],
                    "IGST": r["igst"],
                    "CGST": r["cgst"],
                    "SGST": r["sgst"],
                    "Total Tax": r["total_tax"],
                    "Total Value": r["total_value"],
                    "Source": "GSTR2B",
                    "GSTIN Valid": "Yes" if validate_gstin(r["gstin"]) else "No",
                    "Status": "In GSTR2B, Missing in Tally",
                })
            result.in_gstr2b_not_tally = pd.DataFrame(rows)

        # ---- In GSTR2B but No Bill / No Payment in Tally ----
        # These are GSTR2B entries where:
        #   - Tally has the GSTIN but no matching bill number
        #   - OR Tally has a partial entry (no bill reference / journal only)
        if not unmatched_gstr2b.empty:
            tally_gstins = set(self.tally["gstin"].unique()) - {""}
            rows = []
            for _, r in unmatched_gstr2b.iterrows():
                if r["gstin"] in tally_gstins:
                    # Check if there's any Tally entry with same GSTIN but missing bill
                    tally_same_gstin = self.tally[
                        (self.tally["gstin"] == r["gstin"]) &
                        (~self.tally["_idx"].isin(self._matched_tally_idx))
                    ]
                    has_partial = not tally_same_gstin.empty
                    reason = "Vendor exists in Tally but bill not matched" if has_partial else "GSTIN in Tally but no matching entry"
                else:
                    reason = "Vendor GSTIN not found in Tally at all"

                rows.append({
                    "Supplier Name": r["supplier_name"],
                    "GSTIN": r["gstin"],
                    "Invoice Number": r.get("invoice_number_raw", r["invoice_number"]),
                    "Invoice Date": r.get("invoice_date", ""),
                    "Taxable Value": r["taxable_value"],
                    "IGST": r["igst"],
                    "CGST": r["cgst"],
                    "SGST": r["sgst"],
                    "Total Tax": r["total_tax"],
                    "Total Value": r["total_value"],
                    "Source": "GSTR2B",
                    "Reason": reason,
                    "Status": "No Bill / No Payment in Tally",
                })
            if rows:
                result.no_bill_no_payment = pd.DataFrame(rows)

        # ---- Duplicates ----
        result.duplicates_tally = dup_tally
        result.duplicates_gstr2b = dup_gstr2b

        return result
