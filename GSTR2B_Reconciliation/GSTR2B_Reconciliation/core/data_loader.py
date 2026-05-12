"""
Data loading and cleaning for Tally exports and GSTR2B files.
"""
import logging
import pandas as pd
import numpy as np
from pathlib import Path

from core.utils import (
    clean_gstin, normalize_invoice_number, safe_float, parse_date,
    extract_invoice_from_narration, normalize_name,
)
import config

logger = logging.getLogger("gstr2b_reco")


# ---------------------------------------------------------------------------
# Column Mapping Helper
# ---------------------------------------------------------------------------
def _resolve_columns(df: pd.DataFrame, col_map: dict) -> dict:
    """
    Given a DataFrame and a mapping of {canonical_name: [possible_names]},
    return {canonical_name: actual_column_in_df} for found columns.
    """
    resolved = {}
    df_cols_upper = {c.strip().upper(): c for c in df.columns}
    for canonical, variants in col_map.items():
        for v in variants:
            if v.strip().upper() in df_cols_upper:
                resolved[canonical] = df_cols_upper[v.strip().upper()]
                break
    return resolved


# ---------------------------------------------------------------------------
# Load Tally Data
# ---------------------------------------------------------------------------
def load_tally_data(file_path: str | Path) -> pd.DataFrame:
    """
    Load and clean Tally export (Excel or CSV).
    Returns a DataFrame with standardized column names.
    """
    fp = Path(file_path)
    logger.info(f"Loading Tally data from: {fp}")

    if fp.suffix.lower() in (".xlsx", ".xls"):
        raw = pd.read_excel(fp, dtype=str)
    elif fp.suffix.lower() == ".csv":
        raw = pd.read_csv(fp, dtype=str)
    else:
        raise ValueError(f"Unsupported file format: {fp.suffix}")

    logger.info(f"Tally raw rows: {len(raw)}, columns: {list(raw.columns)}")

    # Resolve columns
    col_map = _resolve_columns(raw, config.TALLY_COLUMN_MAP)
    missing = [k for k in ["party_name", "taxable_value"] if k not in col_map]
    if missing:
        logger.warning(f"Tally: Could not find columns for: {missing}")

    # Build clean DataFrame
    df = pd.DataFrame()
    df["voucher_type"] = raw[col_map["voucher_type"]].str.strip() if "voucher_type" in col_map else "Purchase"
    df["voucher_date_raw"] = raw[col_map["voucher_date"]] if "voucher_date" in col_map else ""
    df["voucher_number"] = raw[col_map["voucher_number"]].str.strip() if "voucher_number" in col_map else ""
    df["party_name"] = raw[col_map["party_name"]].str.strip() if "party_name" in col_map else ""
    df["gstin"] = raw[col_map["gstin"]].apply(clean_gstin) if "gstin" in col_map else ""
    df["bill_number_raw"] = raw[col_map["bill_number"]].str.strip() if "bill_number" in col_map else ""
    df["taxable_value"] = raw[col_map["taxable_value"]].apply(safe_float) if "taxable_value" in col_map else 0.0
    df["igst"] = raw[col_map["igst"]].apply(safe_float) if "igst" in col_map else 0.0
    df["cgst"] = raw[col_map["cgst"]].apply(safe_float) if "cgst" in col_map else 0.0
    df["sgst"] = raw[col_map["sgst"]].apply(safe_float) if "sgst" in col_map else 0.0
    df["total_value"] = raw[col_map["total_value"]].apply(safe_float) if "total_value" in col_map else 0.0
    df["narration"] = raw[col_map["narration"]].fillna("") if "narration" in col_map else ""

    # Parse dates
    df["voucher_date"] = df["voucher_date_raw"].apply(parse_date)

    # Normalize invoice/bill numbers
    df["bill_number"] = df["bill_number_raw"].apply(normalize_invoice_number)

    # Try extracting invoice from narration if bill_number is empty
    mask_empty_bill = df["bill_number"] == ""
    if mask_empty_bill.any():
        df.loc[mask_empty_bill, "bill_number"] = (
            df.loc[mask_empty_bill, "narration"].apply(extract_invoice_from_narration)
        )
        extracted = (df.loc[mask_empty_bill, "bill_number"] != "").sum()
        if extracted:
            logger.info(f"Extracted {extracted} invoice numbers from narrations")

    # Normalize party names for matching
    df["party_name_norm"] = df["party_name"].apply(normalize_name)

    # Compute total tax if missing
    df["total_tax"] = df["igst"] + df["cgst"] + df["sgst"]

    # If total_value is 0, compute it
    mask_no_total = df["total_value"] == 0
    df.loc[mask_no_total, "total_value"] = df.loc[mask_no_total, "taxable_value"] + df.loc[mask_no_total, "total_tax"]

    # Filter by voucher type
    if "voucher_type" in col_map:
        valid_types = [v.upper() for v in config.TALLY_VOUCHER_TYPES]
        df = df[df["voucher_type"].str.upper().isin(valid_types)].copy()

    df["source"] = "Tally"
    df = df.reset_index(drop=True)
    logger.info(f"Tally cleaned rows: {len(df)}")
    return df


# ---------------------------------------------------------------------------
# Load GSTR2B Data
# ---------------------------------------------------------------------------
def load_gstr2b_data(file_path: str | Path) -> pd.DataFrame:
    """
    Load and clean GSTR2B Excel file.
    Returns a DataFrame with standardized column names.
    """
    fp = Path(file_path)
    logger.info(f"Loading GSTR2B data from: {fp}")

    if fp.suffix.lower() in (".xlsx", ".xls"):
        # Pick the B2B sheet (GST portal format) or first sheet
        xl = pd.ExcelFile(fp)
        sheet_names = xl.sheet_names
        target_sheet = next(
            (sn for sn in sheet_names if sn.strip().upper() == "B2B"),
            next((sn for sn in sheet_names if "b2b" in sn.lower()), sheet_names[0]),
        )
        logger.info(f"Using sheet: '{target_sheet}'")

        # Read without headers first so we can detect the structure
        raw_raw = pd.read_excel(fp, sheet_name=target_sheet, header=None, dtype=str)

        # Find the first row containing "GSTIN" — that is the top header row
        header_row_idx = None
        for i in range(min(15, len(raw_raw))):
            row_vals = [str(v).lower() for v in raw_raw.iloc[i].values
                        if pd.notna(v) and str(v).lower() != "nan"]
            if any("gstin" in v for v in row_vals):
                header_row_idx = i
                break

        if header_row_idx is not None:
            # The GST portal format uses TWO header rows (merged cells).
            # Row 1: group labels like "GSTIN of supplier", "Taxable Value (₹)", "Tax Amount" …
            # Row 2: sub-labels like "Invoice number", "Invoice Date", "Integrated Tax(₹)" …
            # For each column we take row-2 value if present, else row-1 value.
            row1 = raw_raw.iloc[header_row_idx].tolist()
            next_idx = header_row_idx + 1
            row2 = raw_raw.iloc[next_idx].tolist() if next_idx < len(raw_raw) else []

            combined = []
            for j, r1 in enumerate(row1):
                r2 = row2[j] if j < len(row2) else None
                r1s = str(r1).strip() if pd.notna(r1) and str(r1) != "nan" else ""
                r2s = str(r2).strip() if pd.notna(r2) and str(r2) != "nan" else ""
                combined.append(r2s if r2s else r1s)

            # Decide data start: if row2 had real sub-headers skip both rows, else skip one
            has_subheaders = any(
                v for v in (raw_raw.iloc[next_idx].tolist() if next_idx < len(raw_raw) else [])
                if pd.notna(v) and str(v).strip() not in ("nan", "")
                and str(v).strip() != str(raw_raw.iloc[header_row_idx][
                    (raw_raw.iloc[next_idx].tolist()).index(v)
                    if v in raw_raw.iloc[next_idx].tolist() else 0
                ]).strip()
            )
            data_start = header_row_idx + (2 if has_subheaders else 1)

            raw = raw_raw.iloc[data_start:].copy()
            raw.columns = combined[: len(raw.columns)]
            raw = raw.reset_index(drop=True)
        else:
            raw = pd.read_excel(fp, sheet_name=target_sheet, dtype=str)

    elif fp.suffix.lower() == ".csv":
        raw = pd.read_csv(fp, dtype=str)
    else:
        raise ValueError(f"Unsupported file format: {fp.suffix}")

    logger.info(f"GSTR2B raw rows: {len(raw)}, columns: {list(raw.columns)}")

    # Resolve columns
    col_map = _resolve_columns(raw, config.GSTR2B_COLUMN_MAP)
    missing = [k for k in ["gstin", "invoice_number", "taxable_value"] if k not in col_map]
    if missing:
        logger.warning(f"GSTR2B: Could not find columns for: {missing}")

    df = pd.DataFrame()
    df["gstin"] = raw[col_map["gstin"]].apply(clean_gstin) if "gstin" in col_map else ""
    df["supplier_name"] = raw[col_map["supplier_name"]].str.strip() if "supplier_name" in col_map else ""
    df["invoice_number_raw"] = raw[col_map["invoice_number"]].str.strip() if "invoice_number" in col_map else ""
    df["invoice_date_raw"] = raw[col_map["invoice_date"]] if "invoice_date" in col_map else ""
    df["taxable_value"] = raw[col_map["taxable_value"]].apply(safe_float) if "taxable_value" in col_map else 0.0
    df["igst"] = raw[col_map["igst"]].apply(safe_float) if "igst" in col_map else 0.0
    df["cgst"] = raw[col_map["cgst"]].apply(safe_float) if "cgst" in col_map else 0.0
    df["sgst"] = raw[col_map["sgst"]].apply(safe_float) if "sgst" in col_map else 0.0
    df["total_tax"] = raw[col_map["total_tax"]].apply(safe_float) if "total_tax" in col_map else 0.0

    # Parse dates
    df["invoice_date"] = df["invoice_date_raw"].apply(parse_date)

    # Normalize invoice numbers
    df["invoice_number"] = df["invoice_number_raw"].apply(normalize_invoice_number)

    # Normalize supplier names
    df["supplier_name_norm"] = df["supplier_name"].apply(normalize_name)

    # Compute total tax if it was 0
    mask_no_tax = df["total_tax"] == 0
    df.loc[mask_no_tax, "total_tax"] = df.loc[mask_no_tax, "igst"] + df.loc[mask_no_tax, "cgst"] + df.loc[mask_no_tax, "sgst"]

    # Compute total value
    df["total_value"] = df["taxable_value"] + df["total_tax"]

    # Drop rows with no GSTIN and no invoice number (likely empty/header rows)
    df = df[~((df["gstin"] == "") & (df["invoice_number"] == ""))].copy()

    df["source"] = "GSTR2B"
    df = df.reset_index(drop=True)
    logger.info(f"GSTR2B cleaned rows: {len(df)}")
    return df
