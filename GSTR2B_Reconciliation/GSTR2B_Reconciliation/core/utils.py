"""
Utility functions for data cleaning, normalization, and validation.
"""
import re
import logging
from datetime import datetime
from difflib import SequenceMatcher

logger = logging.getLogger("gstr2b_reco")

# ---------------------------------------------------------------------------
# GSTIN Validation
# ---------------------------------------------------------------------------
GSTIN_PATTERN = re.compile(
    r"^[0-9]{2}[A-Z]{5}[0-9]{4}[A-Z]{1}[A-Z0-9]{1}Z[A-Z0-9]{1}$"
)


def validate_gstin(gstin: str) -> bool:
    """Return True if the GSTIN matches the standard 15-character format."""
    if not gstin or not isinstance(gstin, str):
        return False
    return bool(GSTIN_PATTERN.match(gstin.strip().upper()))


def clean_gstin(gstin) -> str:
    """Normalize a GSTIN string."""
    if not gstin or (isinstance(gstin, float)):
        return ""
    return str(gstin).strip().upper()


# ---------------------------------------------------------------------------
# Invoice Number Normalization
# ---------------------------------------------------------------------------
_INV_STRIP = re.compile(r"[\s\-/\\\.#'\"]")


def normalize_invoice_number(inv_no) -> str:
    """
    Normalize invoice numbers for comparison:
      - Convert to uppercase
      - Strip whitespace, hyphens, slashes, dots, hash, quotes
      - Remove leading zeros in numeric-only portions
    """
    if not inv_no or (isinstance(inv_no, float)):
        return ""
    s = str(inv_no).strip().upper()
    s = _INV_STRIP.sub("", s)
    # Remove leading zeros only if entire string is numeric
    if s.isdigit():
        s = s.lstrip("0") or "0"
    return s


# ---------------------------------------------------------------------------
# Name Normalization
# ---------------------------------------------------------------------------
_NAME_NOISE = re.compile(r"\b(pvt|private|ltd|limited|llp|inc|co|company)\b", re.I)
_MULTI_SPACE = re.compile(r"\s+")


def normalize_name(name) -> str:
    """Normalize a vendor / party name for fuzzy comparison."""
    if not name or (isinstance(name, float)):
        return ""
    s = str(name).strip().upper()
    s = _NAME_NOISE.sub("", s)
    s = re.sub(r"[^A-Z0-9\s]", "", s)
    s = _MULTI_SPACE.sub(" ", s).strip()
    return s


def fuzzy_name_score(name_a: str, name_b: str) -> int:
    """Return a 0-100 fuzzy similarity score between two names."""
    a = normalize_name(name_a)
    b = normalize_name(name_b)
    if not a or not b:
        return 0
    # Full sequence ratio
    full_ratio = int(SequenceMatcher(None, a, b).ratio() * 100)
    # Token-sorted ratio (sort words alphabetically, then compare)
    a_sorted = " ".join(sorted(a.split()))
    b_sorted = " ".join(sorted(b.split()))
    sorted_ratio = int(SequenceMatcher(None, a_sorted, b_sorted).ratio() * 100)
    # Partial match: check if shorter string is contained in longer
    shorter, longer = (a, b) if len(a) <= len(b) else (b, a)
    partial_ratio = int(SequenceMatcher(None, shorter, longer).ratio() * 100)
    if shorter in longer:
        partial_ratio = max(partial_ratio, 90)
    return max(full_ratio, sorted_ratio, partial_ratio)


# ---------------------------------------------------------------------------
# Amount Helpers
# ---------------------------------------------------------------------------
def safe_float(val, default=0.0) -> float:
    """Convert a value to float safely."""
    if val is None:
        return default
    try:
        return float(val)
    except (ValueError, TypeError):
        return default


def amounts_match(a: float, b: float, tolerance_pct: float = 1.0,
                  tolerance_abs: float = 1.0) -> bool:
    """
    Return True if two amounts are within tolerance.
    Checks both absolute and percentage difference.
    """
    diff = abs(a - b)
    if diff <= tolerance_abs:
        return True
    if a != 0:
        pct_diff = (diff / abs(a)) * 100
        return pct_diff <= tolerance_pct
    return False


# ---------------------------------------------------------------------------
# Date Helpers
# ---------------------------------------------------------------------------
_DATE_FORMATS = [
    "%d-%m-%Y", "%d/%m/%Y", "%Y-%m-%d", "%d-%b-%Y", "%d-%b-%y",
    "%d.%m.%Y", "%m/%d/%Y", "%Y/%m/%d", "%d %b %Y", "%d %B %Y",
    "%d-%m-%y", "%d/%m/%y",
]


def parse_date(val) -> datetime | None:
    """Try multiple date formats and return a datetime or None."""
    if isinstance(val, datetime):
        return val
    if not val or (isinstance(val, float)):
        return None
    s = str(val).strip()
    for fmt in _DATE_FORMATS:
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            continue
    logger.warning(f"Unable to parse date: {s}")
    return None


def dates_within_tolerance(d1, d2, tolerance_days: int = 3) -> bool:
    """Return True if two dates are within ± tolerance_days."""
    if d1 is None or d2 is None:
        return False
    return abs((d1 - d2).days) <= tolerance_days


# ---------------------------------------------------------------------------
# Narration Parsing
# ---------------------------------------------------------------------------
_INV_IN_NARRATION = re.compile(
    r"(?:inv(?:oice)?|bill|ref)[\s.:#\-]*([A-Z0-9/\-]+)", re.I
)


def extract_invoice_from_narration(narration: str) -> str:
    """
    Try to extract an invoice / bill reference from narration text.
    Returns normalized invoice number or empty string.
    """
    if not narration or not isinstance(narration, str):
        return ""
    m = _INV_IN_NARRATION.search(narration)
    if m:
        return normalize_invoice_number(m.group(1))
    return ""
