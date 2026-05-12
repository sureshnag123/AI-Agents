"""
GSTR2B Reconciliation Agent - Streamlit UI

Launch with:
    streamlit run app.py
"""
import sys
import logging
import tempfile
from pathlib import Path
from datetime import datetime
from io import BytesIO

import streamlit as st
import pandas as pd

# Ensure project root is on path
sys.path.insert(0, str(Path(__file__).parent))

import config
from core.data_loader import load_tally_data, load_gstr2b_data
from core.reconciler import Reconciler
from core.report_generator import generate_reports
from core.odoo_connector import OdooConnector

# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------
config.LOG_DIR.mkdir(parents=True, exist_ok=True)
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)-8s | %(message)s",
    handlers=[
        logging.FileHandler(config.LOG_DIR / "app.log", encoding="utf-8"),
        logging.StreamHandler(),
    ],
)
logger = logging.getLogger("gstr2b_reco")

# ---------------------------------------------------------------------------
# Page Config
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title="GSTR2B Reconciliation Agent",
    page_icon="\U0001f9fe",
    layout="wide",
)

# ---------------------------------------------------------------------------
# Custom CSS
# ---------------------------------------------------------------------------
st.markdown("""
<style>
    .main-header {
        font-size: 2rem; font-weight: 700; color: #1B3A5C;
        margin-bottom: 0.5rem;
    }
    .sub-header {
        font-size: 1rem; color: #666; margin-bottom: 2rem;
    }
    .metric-card {
        background: #f7f9fc; border-radius: 8px; padding: 1rem;
        border-left: 4px solid #1B3A5C; margin-bottom: 0.5rem;
    }
    .metric-label { font-size: 0.85rem; color: #666; }
    .metric-value { font-size: 1.4rem; font-weight: 700; color: #1B3A5C; }
    .stTabs [data-baseweb="tab-list"] { gap: 8px; }
    .stTabs [data-baseweb="tab"] {
        background-color: #f0f2f6; border-radius: 4px 4px 0 0; padding: 8px 16px;
    }
    .odoo-badge {
        background: #714B67; color: white; padding: 2px 8px;
        border-radius: 4px; font-size: 0.75rem; font-weight: 600;
    }
    .connected-badge {
        background: #28a745; color: white; padding: 2px 8px;
        border-radius: 4px; font-size: 0.75rem; font-weight: 600;
    }
</style>
""", unsafe_allow_html=True)

# ---------------------------------------------------------------------------
# Header
# ---------------------------------------------------------------------------
st.markdown('<div class="main-header">\U0001f9fe GSTR2B Reconciliation Agent</div>', unsafe_allow_html=True)
st.markdown(
    '<div class="sub-header">Automated ITC reconciliation — Odoo / Tally vs GSTR2B portal data</div>',
    unsafe_allow_html=True,
)

# ---------------------------------------------------------------------------
# Sidebar - Data Source Selection & Controls
# ---------------------------------------------------------------------------
with st.sidebar:
    st.header("\U0001f4c2 Data Source")

    # ── Source A: Purchase data (Odoo or Tally file) ────────────────────────
    source_mode = st.radio(
        "Purchase Data From",
        options=["Odoo (live fetch)", "Tally File (upload)"],
        index=0,
        help="Choose whether to pull invoices from your Odoo instance or upload a Tally export.",
    )

    odoo_df = None   # will be set when user fetches from Odoo

    if source_mode == "Odoo (live fetch)":
        st.markdown('<span class="odoo-badge">Odoo XML-RPC</span>', unsafe_allow_html=True)
        st.markdown("")

        odoo_url = st.text_input(
            "Odoo URL",
            value=st.session_state.get("odoo_url", config.ODOO_URL),
            placeholder="https://mycompany.odoo.com",
        )
        odoo_db = st.text_input(
            "Database",
            value=st.session_state.get("odoo_db", config.ODOO_DB),
            placeholder="mycompany",
        )
        odoo_user = st.text_input(
            "Username / Email",
            value=st.session_state.get("odoo_user", config.ODOO_USERNAME),
            placeholder="admin@mycompany.com",
        )
        odoo_key = st.text_input(
            "API Key / Password",
            type="password",
            value=st.session_state.get("odoo_key", ""),
            help="Use an Odoo API key (Settings → Technical → API Keys) or your login password.",
        )

        fetch_btn = st.button("🔄 Fetch from Odoo", use_container_width=True)

        if fetch_btn:
            if not all([odoo_url, odoo_db, odoo_user, odoo_key]):
                st.error("Please fill in all Odoo connection fields.")
            else:
                with st.spinner("Connecting to Odoo and fetching vendor bills…"):
                    try:
                        conn = OdooConnector(odoo_url, odoo_db, odoo_user, odoo_key)
                        conn.connect()

                        # Save connection params for this session
                        st.session_state["odoo_url"] = odoo_url
                        st.session_state["odoo_db"] = odoo_db
                        st.session_state["odoo_user"] = odoo_user
                        st.session_state["odoo_key"] = odoo_key
                        st.session_state["odoo_conn"] = conn

                        # Fetch bills for the selected month (read from session or default)
                        fetch_month = st.session_state.get("selected_month", "2026-03")
                        df = conn.fetch_vendor_bills(fetch_month)
                        st.session_state["odoo_df"] = df
                        st.session_state["odoo_fetch_month"] = fetch_month

                        st.success(f"Fetched **{len(df)}** vendor bills from Odoo for {fetch_month}")
                    except Exception as e:
                        logger.exception("Odoo fetch failed")
                        st.error(f"Odoo error: {e}")

        if "odoo_df" in st.session_state and not st.session_state["odoo_df"].empty:
            odoo_df = st.session_state["odoo_df"]
            st.markdown(
                f'<span class="connected-badge">✓ {len(odoo_df)} bills loaded</span>',
                unsafe_allow_html=True,
            )

        tally_file = None  # not used in Odoo mode

    else:
        # Tally file upload mode (original behavior)
        tally_file = st.file_uploader(
            "Upload Tally Export",
            type=["xlsx", "xls", "csv"],
            help="Export from Tally Prime: Purchase + Journal vouchers with GST details",
        )

    # ── Source B: GSTR2B file ────────────────────────────────────────────────
    st.divider()
    st.subheader("GSTR2B File")

    gstr2b_input_mode = st.radio(
        "GSTR2B Source",
        options=["Upload file", "Local file path"],
        index=0,
        horizontal=True,
    )

    gstr2b_file = None
    gstr2b_local_path = None

    if gstr2b_input_mode == "Upload file":
        gstr2b_file = st.file_uploader(
            "Upload GSTR2B File",
            type=["xlsx", "xls", "csv"],
            help="GSTR2B Excel file downloaded from the GST Portal",
        )
    else:
        gstr2b_local_path = st.text_input(
            "GSTR2B File Path",
            value=st.session_state.get(
                "gstr2b_path",
                r"C:\Users\User\Downloads\032026_29AACCF2736P1Z5_GSTR2B_15042026.xlsx",
            ),
            placeholder=r"C:\Users\User\Downloads\GSTR2B_file.xlsx",
        )
        if gstr2b_local_path:
            st.session_state["gstr2b_path"] = gstr2b_local_path
            if Path(gstr2b_local_path).exists():
                st.success("File found ✓")
            else:
                st.error("File not found at this path.")

    # ── Settings ─────────────────────────────────────────────────────────────
    st.divider()
    st.header("\u2699\ufe0f Settings")

    month_options = (
        [datetime(2026, m, 1).strftime("%Y-%m") for m in range(1, 13)]
        + [datetime(2025, m, 1).strftime("%Y-%m") for m in range(1, 13)]
    )
    month = st.selectbox(
        "Reconciliation Month",
        options=month_options,
        index=2,   # March 2026
    )
    st.session_state["selected_month"] = month

    with st.expander("Advanced Settings"):
        fuzzy_threshold = st.slider("Fuzzy Match Threshold", 50, 100, config.FUZZY_MATCH_THRESHOLD)
        date_tolerance = st.slider("Date Tolerance (± days)", 0, 15, config.DATE_TOLERANCE_DAYS)
        amt_tol_pct = st.slider("Amount Tolerance (%)", 0.0, 5.0, config.AMOUNT_TOLERANCE_PERCENT, 0.1)
        amt_tol_abs = st.slider("Amount Tolerance (\u20b9)", 0.0, 10.0, config.AMOUNT_TOLERANCE_ABSOLUTE, 0.5)

    st.divider()
    run_btn = st.button("\u25b6\ufe0f  Run Reconciliation", type="primary", use_container_width=True)


# ---------------------------------------------------------------------------
# Readiness Check
# ---------------------------------------------------------------------------
def _gstr2b_ready() -> bool:
    return bool(gstr2b_file) or (gstr2b_local_path and Path(gstr2b_local_path).exists())


def _purchase_ready() -> bool:
    if source_mode == "Odoo (live fetch)":
        return odoo_df is not None and not odoo_df.empty
    return tally_file is not None


if not _purchase_ready() or not _gstr2b_ready():
    if source_mode == "Odoo (live fetch)":
        if not _purchase_ready():
            st.info(
                "**Step 1:** Enter your Odoo credentials in the sidebar and click **Fetch from Odoo** "
                "to load March purchase invoices.\n\n"
                "**Step 2:** Provide the GSTR2B file (upload or local path).\n\n"
                "**Step 3:** Click **Run Reconciliation**."
            )
    else:
        st.info("\U0001f446 Upload both **Tally export** and **GSTR2B file** from the sidebar to get started.")

    with st.expander("\U0001f4cb Odoo — What data is fetched?"):
        st.markdown("""
        The agent connects to Odoo via **XML-RPC** and fetches:
        - All **posted vendor bills** (`in_invoice`) for the selected month
        - All **vendor debit notes** (`in_refund`) for the selected month
        - Fields: Partner name, GSTIN (VAT field on partner), Vendor Reference (bill number),
          Invoice Date, Taxable Value, IGST / CGST / SGST (from journal tax lines), Total Amount

        **Prerequisites in Odoo:**
        - The partner's **VAT / GSTIN** must be filled on the vendor record
        - The **Vendor Reference** (Bill Reference field) must contain the supplier invoice number
        - Tax names must contain **IGST**, **CGST**, or **SGST** / **UTGST** for correct split
        """)

    with st.expander("\U0001f4cb Expected GSTR2B Format"):
        st.markdown("""
        Standard GSTR2B Excel download from the GST Portal.

        | Column | Description |
        |--------|------------|
        | Supplier GSTIN | 15-character GSTIN |
        | Invoice Number | Supplier invoice reference |
        | Invoice Date | Date of invoice |
        | Taxable Value | Base amount |
        | IGST / CGST / SGST | Tax components |
        """)

    st.stop()


# ---------------------------------------------------------------------------
# Run Reconciliation
# ---------------------------------------------------------------------------
if run_btn:
    with st.spinner("Running reconciliation…"):
        try:
            # ── Load purchase-side data ───────────────────────────────────
            if source_mode == "Odoo (live fetch)":
                purchase_df = odoo_df.copy()
                logger.info(f"Using Odoo data: {len(purchase_df)} records")
            else:
                with tempfile.NamedTemporaryFile(
                    suffix=f".{tally_file.name.split('.')[-1]}", delete=False
                ) as tf:
                    tf.write(tally_file.getbuffer())
                    tally_path = tf.name
                purchase_df = load_tally_data(tally_path)

            # ── Load GSTR2B data ──────────────────────────────────────────
            if gstr2b_file:
                with tempfile.NamedTemporaryFile(
                    suffix=f".{gstr2b_file.name.split('.')[-1]}", delete=False
                ) as gf:
                    gf.write(gstr2b_file.getbuffer())
                    gstr2b_path = gf.name
            else:
                gstr2b_path = gstr2b_local_path

            gstr2b_df = load_gstr2b_data(gstr2b_path)

            # ── Run reconciliation engine ─────────────────────────────────
            reconciler = Reconciler(
                purchase_df, gstr2b_df,
                fuzzy_threshold=fuzzy_threshold,
                date_tolerance=date_tolerance,
                amt_tolerance_pct=amt_tol_pct,
                amt_tolerance_abs=amt_tol_abs,
            )
            result = reconciler.run()

            # ── Generate Excel reports ────────────────────────────────────
            output_dir = config.OUTPUT_DIR / month
            files = generate_reports(result, output_dir, month)

            st.session_state["result"] = result
            st.session_state["files"] = files
            st.session_state["month"] = month
            st.session_state["source_label"] = (
                "Odoo" if source_mode == "Odoo (live fetch)" else "Tally"
            )

            st.success(f"Reconciliation complete in {result.run_time_seconds:.2f} seconds!")

        except Exception as e:
            logger.exception("Reconciliation failed")
            st.error(f"Error: {e}")


# ---------------------------------------------------------------------------
# Display Results
# ---------------------------------------------------------------------------
if "result" in st.session_state:
    result = st.session_state["result"]
    files = st.session_state["files"]
    s = result.summary
    source_label = st.session_state.get("source_label", "Purchase")

    # Metrics row
    cols = st.columns(6)
    metric_data = [
        ("\u2705 Matched", s["matched_count"]),
        ("\u26a0\ufe0f Amt Mismatch", s["amount_mismatch"]),
        (f"\U0001f4d5 {source_label} Only", s["in_tally_not_gstr2b"]),
        ("\U0001f4d9 GSTR2B Only", s["in_gstr2b_not_tally"]),
        ("\U0001f6ab No Bill/Pay", s["no_bill_no_payment"]),
        ("\u23f1\ufe0f Time (sec)", s["run_time_seconds"]),
    ]
    for col, (label, value) in zip(cols, metric_data):
        col.metric(label, value)

    # ITC comparison
    st.divider()
    c1, c2, c3 = st.columns(3)
    c1.metric(f"ITC as per {source_label} (\u20b9)", f"{s['tally_itc_total']:,.2f}")
    c2.metric("ITC as per GSTR2B (\u20b9)", f"{s['gstr2b_itc_total']:,.2f}")
    diff = s["tally_itc_total"] - s["gstr2b_itc_total"]
    c3.metric("ITC Difference (\u20b9)", f"{diff:,.2f}", delta=f"{diff:,.2f}")

    st.divider()

    # Tabs for detailed views
    tabs = st.tabs([
        "\u2705 Matched",
        f"\U0001f4d5 In {source_label}, NOT GSTR2B",
        "\U0001f4d9 In GSTR2B, NOT " + source_label,
        "\U0001f6ab No Bill/Payment",
        "\u26a0\ufe0f Amount Mismatch",
        "\U0001f503 Duplicates",
    ])

    with tabs[0]:
        if result.matched is not None and not result.matched.empty:
            st.dataframe(result.matched, use_container_width=True, hide_index=True)
        else:
            st.info("No matched entries found.")

    with tabs[1]:
        if result.in_tally_not_gstr2b is not None and not result.in_tally_not_gstr2b.empty:
            st.dataframe(result.in_tally_not_gstr2b, use_container_width=True, hide_index=True)
        else:
            st.info(f"All {source_label} entries matched with GSTR2B.")

    with tabs[2]:
        if result.in_gstr2b_not_tally is not None and not result.in_gstr2b_not_tally.empty:
            st.dataframe(result.in_gstr2b_not_tally, use_container_width=True, hide_index=True)
        else:
            st.info(f"All GSTR2B entries matched with {source_label}.")

    with tabs[3]:
        if result.no_bill_no_payment is not None and not result.no_bill_no_payment.empty:
            st.dataframe(result.no_bill_no_payment, use_container_width=True, hide_index=True)
        else:
            st.info("No entries in this category.")

    with tabs[4]:
        if result.amount_mismatch is not None and not result.amount_mismatch.empty:
            st.dataframe(result.amount_mismatch, use_container_width=True, hide_index=True)
        else:
            st.info("No amount mismatches found.")

    with tabs[5]:
        has_dup_t = result.duplicates_tally is not None and not result.duplicates_tally.empty
        has_dup_g = result.duplicates_gstr2b is not None and not result.duplicates_gstr2b.empty
        if has_dup_t or has_dup_g:
            if has_dup_t:
                st.subheader(f"{source_label} Duplicates")
                st.dataframe(result.duplicates_tally, use_container_width=True, hide_index=True)
            if has_dup_g:
                st.subheader("GSTR2B Duplicates")
                st.dataframe(result.duplicates_gstr2b, use_container_width=True, hide_index=True)
        else:
            st.info("No duplicate entries detected.")

    # Download buttons
    st.divider()
    st.subheader("\U0001f4e5 Download Reports")
    dl_cols = st.columns(2)

    for i, (name, path) in enumerate(files.items()):
        with open(path, "rb") as f:
            dl_cols[i % 2].download_button(
                label=f"\U0001f4c4 {Path(path).name}",
                data=f.read(),
                file_name=Path(path).name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
