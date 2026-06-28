"""
Koricube Internal Operations Console — "Kori Frost" Edition
===========================================================
A Streamlit front-end for the "Koricube" automated ice-machine business,
styled as a clean, minimalist, crisp SaaS dashboard (Stripe / Vercel feel).

The single source of truth is a Google Sheets workbook named ``Koricube_Database``
containing four worksheets:

    - Location      (master data: machines, branches, payment config)   [READ]
    - Sales_Log     (daily cash reconciliation)                         [APPEND]
    - Raw_Email     (parsed bank settlement rows)                       [APPEND]
    - Maintenance   (repair tickets)                                    [APPEND]

CRITICAL ARCHITECTURE RULES (UNBREAKABLE)
-----------------------------------------
1. Row 1 of every *write* sheet holds live ArrayFormula / MAP+LAMBDA logic.
   We therefore NEVER touch row 1 and exclusively use ``worksheet.append_row()``
   which writes to the first fully-empty row at the bottom of the sheet.
2. Every date pushed to Google Sheets is normalised to a strict ``YYYY-MM-DD``
   string so the upstream formulas never misparse a locale-specific date.
3. All timestamps/dates are generated in the ``Asia/Bangkok`` timezone.

Theming note
------------
The base palette (ice-blue primary, off-white canvas, slate text) lives in
``.streamlit/config.toml`` so native widgets — sidebar, selects, number-input
steppers — keep their correct light styling. The CSS below ONLY layers
typography + card/badge surfaces on top; it deliberately does not restyle
inputs, steppers, or the sidebar.

Author: Koricube Engineering
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import datetime, date
from html import escape
from typing import Any, List, Optional

import altair as alt
import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

import gspread

# ---------------------------------------------------------------------------
# Timezone handling (Asia/Bangkok) with a graceful fallback chain.
# ---------------------------------------------------------------------------
try:
    from zoneinfo import ZoneInfo  # Python 3.9+
    BKK_TZ = ZoneInfo("Asia/Bangkok")
except Exception:  # pragma: no cover - very old runtimes / missing tzdata
    try:
        import pytz
        BKK_TZ = pytz.timezone("Asia/Bangkok")
    except Exception:
        BKK_TZ = None  # last resort: naive local time


# ===========================================================================
# CONSTANTS
# ===========================================================================
SERVICE_ACCOUNT_FILE = "service_account.json"
SPREADSHEET_NAME = "Koricube_Database"

# Worksheet (tab) names.
WS_LOCATION = "Location"
WS_SALES_LOG = "Sales_Log"
WS_RAW_EMAIL = "Raw_Email"
WS_MAINTENANCE = "Maintenance"
WS_FINAL_DB = "Final_Database"   # legacy, pre-calculated history for the dashboard

# Location column that stores each machine's fixed monthly rent (dashboard only).
LOC_RENT_COST = "Rent_Cost"

# Location master-data column headers (row 1 of the Location sheet).
LOC_MACHINE_ID = "Machine_ID"
LOC_BRANCH = "Current_Branch"
LOC_MACHINE_TYPE = "Machine_Type"
LOC_PAYMENT_TYPE = "Payment_Type"
LOC_MERCHANT_NO = "Merchant_No"
LOC_MACHINES_SHARED = "Machines_Shared"

LOCATION_COLUMNS = [
    LOC_MACHINE_ID,
    LOC_BRANCH,
    LOC_MACHINE_TYPE,
    LOC_PAYMENT_TYPE,
    LOC_MERCHANT_NO,
    LOC_MACHINES_SHARED,
]

# Payment-type business values (exact Thai strings as stored in the sheet).
PAYMENT_CASH_ONLY = "เงินสดเท่านั้น"      # Cash only -> hide period fields
PAYMENT_TRANSFER_CASH = "โอน+เงินสด"      # Transfer + cash -> show period fields

# Maintenance & Expense ledger codes.
# The dropdown VALUE stored in the sheet's Error_Code column is the raw code
# (e.g. "BILL-WATER"); the Thai gloss is display-only (see MAINTENANCE_CODE_LABELS).
ERROR_CODES = ["E7", "A05", "Other"]                          # technical repairs
EXPENSE_CODES = ["BILL-WATER", "BILL-ELEC", "MISC-EXPENSE"]   # variable expenses
MAINTENANCE_CODES = ERROR_CODES + EXPENSE_CODES
MAINTENANCE_CODE_LABELS = {
    "E7": "E7",
    "A05": "A05",
    "Other": "Other (อื่นๆ)",
    "BILL-WATER": "BILL-WATER (ค่าน้ำ)",
    "BILL-ELEC": "BILL-ELEC (ค่าไฟ)",
    "MISC-EXPENSE": "MISC-EXPENSE (ค่าใช้จ่ายอื่นๆ)",
}

STATUS_PENDING = "รอซ่อม"          # default status for a technical repair ticket
STATUS_EXPENSE_PAID = "เคลียร์แล้ว"  # default status for a settled expense entry

# Maintenance sheet schema. Branch_Name (col C) is STAMPED from Location at
# write time — mirrors the Sales_Log layout so the dashboard maps both alike.
MAINTENANCE_COLUMNS = [
    "Report_Date",    # A
    "Machine_ID",     # B
    "Branch_Name",    # C  (stamped)
    "Error_Code",     # D
    "Issue_Desc",     # E
    "Repair_Cost",    # F
    "Resolved_Date",  # G
    "Status",         # H
]

# Buddhist Era offset (BE = AD + 543).
BE_OFFSET = 543

# ``USER_ENTERED`` lets Google Sheets store our numeric strings as real numbers
# and our unambiguous ISO ``YYYY-MM-DD`` strings as real dates, so the row-1
# ArrayFormula logic can compute on them directly.
VALUE_INPUT_OPTION = "USER_ENTERED"

# Sales_Log full column schema (A → P). The first 11 are the original entry
# fields; the trailing five (L→P) are the reconciliation outputs now written
# explicitly from Python for full transparency.
SALES_LOG_COLUMNS = [
    "Collection_Date",    # A
    "Machine_ID",         # B
    "Branch_Name",        # C
    "Merchant_No",        # D
    "Payment_Type",       # E
    "Machines_Shared",    # F
    "Web_Total",          # G
    "Cash_Collected",     # H
    "Period_Start",       # I
    "Period_End",         # J
    "Remark",             # K
    "Expected_Transfer",  # L
    "Shared_Bank_Fee",    # M
    "Net_Actual",         # N
    "Status",             # O
    "Diff_Adjustment",    # P
]

# Status values stamped into the Sales_Log "Status" column (O).
RECON_STATUS_DONE = "กระทบยอดแล้ว"   # reconciled against bank data (Reconciled)
RECON_STATUS_CASH = "เงินสด"          # cash-only machine, no bank transfer to reconcile


# ===========================================================================
# TIME / DATE HELPERS
# ===========================================================================
def now_bkk() -> datetime:
    """Return the current datetime in the Asia/Bangkok timezone."""
    if BKK_TZ is not None:
        return datetime.now(BKK_TZ)
    return datetime.now()


def today_iso() -> str:
    """Today's date in Bangkok as a strict ``YYYY-MM-DD`` string."""
    return now_bkk().strftime("%Y-%m-%d")


def date_to_iso(value: Optional[date]) -> str:
    """Format a ``date``/``datetime`` as ``YYYY-MM-DD`` (or '' when None)."""
    if value is None:
        return ""
    if isinstance(value, datetime):
        value = value.date()
    return value.strftime("%Y-%m-%d")


def convert_be_year_to_ad(year: int) -> int:
    """
    Convert a Buddhist-Era year to the Gregorian (AD) calendar.

    Any 4-digit year >= 2400 is treated as BE and reduced by 543
    (e.g. 2569 -> 2026). Years already in AD pass through untouched.
    """
    return year - BE_OFFSET if year >= 2400 else year


def normalize_date_string(raw: str) -> str:
    """
    Best-effort normalisation of a free-form date token into ``YYYY-MM-DD``.

    Handles the common shapes seen in Thai bank PDFs and automatically
    downgrades Buddhist-Era years to AD:

        "23/06/2569" -> "2026-06-23"
        "2569-06-23" -> "2026-06-23"
        "23-06-2026" -> "2026-06-23"

    Returns the original string untouched if no known pattern matches, so the
    operator can still review/fix it before committing.
    """
    if raw is None:
        return ""
    token = str(raw).strip()
    if not token:
        return ""

    # Pull the first 3 integer groups out of the token.
    parts = re.findall(r"\d+", token)
    if len(parts) < 3:
        return token  # not a recognisable date -> leave for manual review

    a, b, c = parts[0], parts[1], parts[2]

    # Decide which group is the year (the 4-digit one).
    if len(a) == 4:                      # YYYY-MM-DD style
        year, month, day = int(a), int(b), int(c)
    elif len(c) == 4:                    # DD-MM-YYYY style
        day, month, year = int(a), int(b), int(c)
    else:
        return token                     # ambiguous -> manual review

    year = convert_be_year_to_ad(year)

    try:
        return date(year, month, day).strftime("%Y-%m-%d")
    except ValueError:
        return token                     # invalid calendar date -> leave as-is


def clean_number(raw: Any) -> Optional[float]:
    """
    Coerce a messy numeric string ("1,234.50 ฿", "(50.00)") into a float.

    Returns ``None`` when the value cannot be parsed as a number.
    """
    if raw is None:
        return None
    if isinstance(raw, (int, float)):
        return float(raw)

    text = str(raw).strip()
    if not text:
        return None

    # Accounting-style negatives in parentheses.
    negative = text.startswith("(") and text.endswith(")")

    # Strip everything except digits, separators and sign.
    text = re.sub(r"[^0-9.,\-]", "", text)
    text = text.replace(",", "")

    if text in ("", "-", ".", "-."):
        return None
    try:
        number = float(text)
    except ValueError:
        return None
    return -number if negative else number


# ===========================================================================
# GOOGLE SHEETS CONNECTION LAYER
# ===========================================================================
@st.cache_resource(show_spinner="Connecting to Google Sheets…")
def get_client() -> gspread.Client:
    """
    Authenticate once and cache the gspread client for the whole session.

    Deploy-friendly: on Streamlit Cloud the key lives in st.secrets under
    ``[gcp_service_account]`` (nothing committed to git); locally it falls back
    to the ``service_account.json`` file for development.
    """
    try:
        secret = st.secrets["gcp_service_account"]
    except Exception:  # noqa: BLE001 - no secrets file locally -> use the JSON
        secret = None
    if secret:
        return gspread.service_account_from_dict(dict(secret))
    return gspread.service_account(filename=SERVICE_ACCOUNT_FILE)


@st.cache_resource(show_spinner="Opening Koricube_Database…")
def get_spreadsheet() -> gspread.Spreadsheet:
    """Open (and cache) the ``Koricube_Database`` spreadsheet handle."""
    return get_client().open(SPREADSHEET_NAME)


def get_worksheet(name: str) -> gspread.Worksheet:
    """Return a worksheet handle by tab name."""
    return get_spreadsheet().worksheet(name)


@st.cache_data(ttl=300, show_spinner="Loading machine master data…")
def fetch_location_data() -> pd.DataFrame:
    """
    Read the Location master sheet into a tidy DataFrame.

    Cached for 5 minutes; use the sidebar "Refresh data" button to force a
    re-read after the master sheet changes. All values are returned as strings
    to avoid surprises with merchant numbers / IDs that look numeric.
    """
    worksheet = get_worksheet(WS_LOCATION)
    # ``expected_headers`` keeps get_all_records robust if extra helper columns
    # exist in the sheet beyond the six we care about.
    records = worksheet.get_all_records(expected_headers=LOCATION_COLUMNS)
    df = pd.DataFrame(records)

    # Guarantee every expected column exists even on a partially-filled sheet.
    for col in LOCATION_COLUMNS:
        if col not in df.columns:
            df[col] = ""

    df = df[LOCATION_COLUMNS].astype(str).apply(lambda s: s.str.strip())
    # Drop blank rows (no machine id).
    df = df[df[LOC_MACHINE_ID] != ""].reset_index(drop=True)
    return df


def append_row_safe(worksheet_name: str, payload: List[Any]) -> None:
    """
    Append a single row to the bottom of a worksheet.

    ``append_row`` always targets the first empty row, so row 1 (ArrayFormula)
    is never disturbed.
    """
    worksheet = get_worksheet(worksheet_name)
    worksheet.append_row(payload, value_input_option=VALUE_INPUT_OPTION)


# ===========================================================================
# RAW_EMAIL SCHEMA  (bank settlement rows — populated by the statement bot)
# ===========================================================================
# Column order of the Raw_Email sheet, consumed by the date-range aggregation
# (fetch_raw_email_data / aggregate_bank_transfers). The in-app PDF parser has
# been retired now that a bot writes these rows directly.
RAW_EMAIL_FIELDS = [
    "Transfer_Date",
    "Merchant_No",
    "Trans_Amount",
    "Commission",
    "VAT",
    "Net_Transfer",
]


# ===========================================================================
# SHARED-MERCHANT RECONCILIATION LOGIC
# ===========================================================================
# Scenario: one Ksher merchant / one bank transfer covers TWO machines
# (Machines_Shared == 2). The single transfer and its fee must be split fairly
# between the two machines. This mirrors the finalised Excel model 1:1.
#
#   Expected_X      = Web_Total_X - Cash_Collected_X          (per machine)
#   Total_Expected  = Expected_A + Expected_B
#   Diff            = Actual_Transfer(X) - Total_Expected
#   Diff_Adjustment = Diff / shares                           (split evenly)
#   Shared_Bank_Fee = Total_Fee(Y) / shares                  (split evenly)
#   Net_Actual_X    = Expected_X + Diff_Adjustment - Shared_Bank_Fee
#
# By construction the stored columns stay internally consistent:
#   Net_Actual == Expected_Transfer + Diff_Adjustment - Shared_Bank_Fee


def _to_float(value: Any, default: float = 0.0) -> float:
    """Robustly coerce any cell/input into a float (blank/garbage -> default)."""
    parsed = clean_number(value)
    return float(parsed) if parsed is not None else float(default)


@dataclass(frozen=True)
class SharedReconResult:
    """Immutable result of a 2-machine shared-merchant reconciliation."""
    expected_a: float
    expected_b: float
    total_expected: float
    actual_transfer: float    # X
    diff: float               # X - total_expected
    diff_adjustment: float    # diff / shares
    total_fee: float          # Y
    shared_bank_fee: float    # Y / shares
    net_actual_a: float
    net_actual_b: float


def compute_shared_reconciliation(
    web_a: Any, cash_a: Any,
    web_b: Any, cash_b: Any,
    actual_transfer: Any, total_fee: Any,
    shares: int = 2,
) -> SharedReconResult:
    """
    Translate the finalised Excel reconciliation into Python (see section header).

    Verified 1:1 against the Excel model (e.g. row 3: Expected 1944/1043,
    Diff 43, Diff/2 21.5, Fee/2 9.95 -> Net 1955.55 / 1054.55).

    All inputs are coerced defensively so blank/None/"1,234.50" never crash.
    ``shares`` is the divisor for the fair split (2 for this scenario); it is
    floored to 1 to guarantee no division-by-zero. Results are rounded to 2 dp
    AND the net is derived from the rounded split components, so the stored
    columns always reconcile (Net == Expected + Diff_Adjustment - Shared_Fee).
    """
    shares = max(int(shares or 1), 1)

    web_a, cash_a = _to_float(web_a), _to_float(cash_a)
    web_b, cash_b = _to_float(web_b), _to_float(cash_b)
    x, y = _to_float(actual_transfer), _to_float(total_fee)

    expected_a = round(web_a - cash_a, 2)
    expected_b = round(web_b - cash_b, 2)
    total_expected = round(expected_a + expected_b, 2)

    diff = round(x - total_expected, 2)
    diff_adjustment = round(diff / shares, 2)
    shared_bank_fee = round(y / shares, 2)

    net_actual_a = round(expected_a + diff_adjustment - shared_bank_fee, 2)
    net_actual_b = round(expected_b + diff_adjustment - shared_bank_fee, 2)

    return SharedReconResult(
        expected_a=expected_a, expected_b=expected_b,
        total_expected=total_expected, actual_transfer=round(x, 2),
        diff=diff, diff_adjustment=diff_adjustment,
        total_fee=round(y, 2), shared_bank_fee=shared_bank_fee,
        net_actual_a=net_actual_a, net_actual_b=net_actual_b,
    )


def build_shared_reconciliation_rows(
    *, collection_date: str, period_start: str, period_end: str, remark: str,
    machine_a: pd.Series, machine_b: pd.Series,
    web_a: Any, cash_a: Any, web_b: Any, cash_b: Any,
    result: SharedReconResult, status: str = RECON_STATUS_DONE,
) -> List[List[Any]]:
    """
    Prepare the two append-ready Sales_Log rows (full 16-column schema, A→P).

    Master fields (Branch / Merchant / Payment / Shared) are stamped from the
    CURRENT Location data, exactly like the single-machine flow. Returns
    ``[row_a, row_b]``; the caller appends each with ``append_row_safe`` so
    row 1 (ArrayFormula) is never touched.
    """
    def _row(machine: pd.Series, web: Any, cash: Any,
             expected: float, net: float) -> List[Any]:
        return [
            collection_date,               # A  Collection_Date
            machine[LOC_MACHINE_ID],       # B  Machine_ID
            machine[LOC_BRANCH],           # C  Branch_Name
            machine[LOC_MERCHANT_NO],      # D  Merchant_No
            machine[LOC_PAYMENT_TYPE],     # E  Payment_Type
            machine[LOC_MACHINES_SHARED],  # F  Machines_Shared
            _to_float(web),                # G  Web_Total
            _to_float(cash),               # H  Cash_Collected
            period_start,                  # I  Period_Start
            period_end,                    # J  Period_End
            remark,                        # K  Remark
            expected,                      # L  Expected_Transfer
            result.shared_bank_fee,        # M  Shared_Bank_Fee
            net,                           # N  Net_Actual
            status,                        # O  Status
            result.diff_adjustment,        # P  Diff_Adjustment
        ]

    return [
        _row(machine_a, web_a, cash_a, result.expected_a, result.net_actual_a),
        _row(machine_b, web_b, cash_b, result.expected_b, result.net_actual_b),
    ]


# ---------------------------------------------------------------------------
# Raw_Email auto-aggregation — sum the actual transfer (X) and fee (Y) for a
# merchant across a Period_Start..Period_End date range.
# ---------------------------------------------------------------------------
@st.cache_data(ttl=120, show_spinner="Loading bank transfer data…")
def fetch_raw_email_data() -> pd.DataFrame:
    """
    Read the Raw_Email settlement sheet into a tidy DataFrame (cached 2 min).

    Use the sidebar "Refresh" button to force a re-read after new bank
    statements are appended.
    """
    worksheet = get_worksheet(WS_RAW_EMAIL)
    records = worksheet.get_all_records(expected_headers=RAW_EMAIL_FIELDS)
    df = pd.DataFrame(records)
    for col in RAW_EMAIL_FIELDS:
        if col not in df.columns:
            df[col] = ""
    return df[RAW_EMAIL_FIELDS]


def _digits_only(value: Any) -> str:
    """Normalise a merchant number to digits only for robust matching."""
    return re.sub(r"\D", "", str(value or ""))


def aggregate_bank_transfers(
    raw_df: pd.DataFrame, merchant_no: Any, start_iso: str, end_iso: str
) -> Optional[tuple]:
    """
    Sum every Raw_Email settlement for ``merchant_no`` whose Transfer_Date falls
    within the inclusive range [start, end].

    Returns ``(X, Y)`` where:
        X = SUM(Trans_Amount)   -> GROSS Actual Total Transfer
        Y = SUM(Commission)     -> Total Bank Fee
    rounded to 2 dp. Returns ``None`` when the merchant/dates are missing, no
    rows fall in range, OR the transfer total X is 0 (treated as "not in yet").

    Robustness: dates are normalised (BE->AD aware) and parsed with
    ``errors='coerce'`` so blank/garbage Transfer_Date values become NaT and are
    silently excluded — empty sets never raise.
    """
    target_merchant = _digits_only(merchant_no)
    # A single supplied bound collapses the range to that one day.
    start_iso = (start_iso or end_iso or "").strip()
    end_iso = (end_iso or start_iso or "").strip()
    if raw_df is None or raw_df.empty or not target_merchant or not start_iso:
        return None

    start_dt = pd.to_datetime(start_iso, errors="coerce")
    end_dt = pd.to_datetime(end_iso, errors="coerce")
    if pd.isna(start_dt) or pd.isna(end_dt):
        return None

    df = raw_df.copy()
    df["_m"] = df["Merchant_No"].apply(_digits_only)
    df["_d"] = pd.to_datetime(
        df["Transfer_Date"].apply(normalize_date_string), errors="coerce"
    )
    mask = (df["_m"] == target_merchant) & (df["_d"] >= start_dt) & (df["_d"] <= end_dt)
    match = df[mask]
    if match.empty:
        return None

    x = round(float(match["Trans_Amount"].apply(_to_float).sum()), 2)
    y = round(float(match["Commission"].apply(_to_float).sum()), 2)
    if x == 0:  # no real transfer landed in this window yet
        return None
    return x, y


def safe_aggregate_bank_transfers(
    merchant_no: Any, start_iso: str, end_iso: str
) -> Optional[tuple]:
    """Cached Raw_Email aggregation that degrades gracefully if the sheet is unreadable."""
    try:
        raw_df = fetch_raw_email_data()
    except Exception:  # noqa: BLE001 - treat a read failure as "no data yet"
        return None
    return aggregate_bank_transfers(raw_df, merchant_no, start_iso, end_iso)


# ===========================================================================
# PRESENTATION LAYER — "Kori Frost" CSS / HTML COMPONENTS (cosmetic only)
# ===========================================================================
def inject_css() -> None:
    """
    Layer the Kori Frost surfaces on top of the config.toml base theme.

    Scope is intentionally narrow:
      - typography (Prompt / Inter)
      - the minimalist top bar (NO gradient)
      - section headings, the compact status badge, the reconciliation cue
      - soft Kori shadow on cards (forms + bordered containers)
      - a light hover polish on PRIMARY buttons only

    It deliberately does NOT target inputs, number steppers, selects, or the
    sidebar, so Streamlit's native light theme renders them untouched.
    """
    st.markdown(
        """
        <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=Prompt:wght@400;500;600&display=swap');

        :root{
            --kc-bg:#F8FAFC; --kc-card:#FFFFFF; --kc-border:#E2E8F0;
            --kc-ink:#0F172A; --kc-muted:#64748B;
            --kc-accent:#0EA5E9; --kc-accent-strong:#0284C7;
            --kc-accent-soft:#E0F2FE; --kc-accent-line:#BAE6FD;
            --kc-shadow:0 4px 6px -1px rgba(0,0,0,.05);
            --kc-radius:14px;
        }

        html, body, [class*="css"], p, span, label, div, button{
            font-family:'Inter','Prompt',system-ui,sans-serif;
        }
        h1,h2,h3,h4{
            font-family:'Prompt','Inter',sans-serif;
            color:var(--kc-ink); letter-spacing:-.01em;
        }
        /* ---- Responsive content width (sidebar-toggle fix) ----
           Fluid + ALWAYS centered, whether the sidebar is open or collapsed.
           1) Zero the collapsed sidebar so it stops reserving ~340px on the
              left (the root cause of the "pushed right / gap on left" bug).
           2) Let the main area span the full available width.
           3) A generous cap + forced auto side-margins keep content centered
              and let it breathe instead of being locked to a narrow column. */
        section[data-testid="stSidebar"][aria-expanded="false"]{
            width:0 !important; min-width:0 !important; margin-left:0 !important;
        }
        [data-testid="stMain"]{ width:100% !important; }
        /* Keep Streamlit's fixed top strip (where the >> toggle lives) on the
           Kori Frost canvas colour and above the content, so nothing can ever
           show through behind the toggle button. */
        [data-testid="stHeader"]{
            background:var(--kc-bg) !important;
            height:3.5rem !important; z-index:1000;
        }
        .block-container, [data-testid="stMainBlockContainer"]{
            max-width:1500px !important;
            margin-left:auto !important; margin-right:auto !important;
            /* Fixed, generous top padding ALWAYS clears the header — no overlap
               with the >> button whether the sidebar is open or collapsed.
               No negative margins / absolute positioning anywhere. */
            padding-top:4.5rem !important; padding-bottom:3rem !important;
            padding-left:clamp(1rem,4vw,3rem) !important;
            padding-right:clamp(1rem,4vw,3rem) !important;
        }
        #MainMenu, footer{ visibility:hidden; }

        /* ---- Minimalist top bar (no gradient) ---- */
        .kc-topbar{ display:flex; align-items:center; gap:14px; margin:2px 0 20px; }
        .kc-logo{
            width:44px; height:44px; border-radius:12px; display:grid; place-items:center;
            background:var(--kc-accent-soft); color:var(--kc-accent-strong);
            font-size:1.4rem; border:1px solid var(--kc-accent-line);
        }
        .kc-topbar .t{ font-family:'Prompt',sans-serif; font-weight:600; font-size:1.3rem; color:var(--kc-ink); line-height:1.15; }
        .kc-topbar .s{ color:var(--kc-muted); font-size:.84rem; margin-top:2px; }
        .kc-topbar .clock{
            margin-left:auto; font-size:.8rem; color:var(--kc-muted);
            background:var(--kc-card); border:1px solid var(--kc-border);
            border-radius:999px; padding:6px 14px; box-shadow:var(--kc-shadow); white-space:nowrap;
        }

        /* ---- Section heading ---- */
        .kc-sec{ display:flex; align-items:center; gap:11px; margin:2px 0 12px; }
        .kc-sec .i{
            width:34px; height:34px; border-radius:9px; display:grid; place-items:center;
            background:var(--kc-accent-soft); color:var(--kc-accent-strong);
            font-size:1rem; border:1px solid var(--kc-accent-line);
        }
        .kc-sec .h{ font-family:'Prompt',sans-serif; font-weight:600; font-size:1.02rem; color:var(--kc-ink); }
        .kc-sec .d{ font-size:.8rem; color:var(--kc-muted); margin-top:1px; }

        /* ---- Compact single-line status badge row ---- */
        .kc-status{
            display:flex; flex-wrap:wrap; align-items:center; gap:10px 22px;
            background:var(--kc-card); border:1px solid var(--kc-border);
            border-radius:var(--kc-radius); padding:12px 18px;
            box-shadow:var(--kc-shadow); margin-bottom:6px;
        }
        .kc-status .it{ display:flex; align-items:baseline; gap:7px; }
        .kc-status .it .k{ font-size:.68rem; text-transform:uppercase; letter-spacing:.05em; color:var(--kc-muted); font-weight:600; }
        .kc-status .it .v{ font-size:.92rem; font-weight:600; color:var(--kc-ink); }
        .kc-status .sep{ width:1px; height:18px; background:var(--kc-border); }
        .kc-pill{
            display:inline-flex; align-items:center; gap:6px; margin-left:auto;
            padding:5px 12px; border-radius:999px; font-size:.82rem; font-weight:600;
        }
        .kc-pill.transfer{ background:var(--kc-accent-soft); color:var(--kc-accent-strong); border:1px solid var(--kc-accent-line); }
        .kc-pill.cash{ background:#F1F5F9; color:#475569; border:1px solid var(--kc-border); }

        /* ---- Reconciliation cue (subtle, crisp) ---- */
        .kc-rec{
            display:flex; align-items:center; justify-content:space-between; gap:10px;
            padding:10px 16px; border-radius:12px; font-weight:600;
            margin-top:2px; border:1px solid var(--kc-border); background:var(--kc-card);
        }
        .kc-rec .lbl{ font-size:.8rem; font-weight:600; color:var(--kc-muted); }
        .kc-rec .val{ font-size:1rem; font-weight:700; }
        .kc-rec.ok{    background:#F0FDF4; border-color:#BBF7D0; } .kc-rec.ok .val{    color:#15803D; }
        .kc-rec.over{  background:#FFF7ED; border-color:#FED7AA; } .kc-rec.over .val{  color:#C2410C; }
        .kc-rec.short{ background:#FEF2F2; border-color:#FECACA; } .kc-rec.short .val{ color:#B91C1C; }

        /* ---- Cards: forms + bordered containers get the soft Kori shadow ---- */
        [data-testid="stForm"]{
            background:var(--kc-card); border:1px solid var(--kc-border);
            border-radius:16px; padding:18px 20px; box-shadow:var(--kc-shadow);
        }
        [data-testid="stVerticalBlockBorderWrapper"]{
            border-radius:16px; box-shadow:var(--kc-shadow);
        }

        /* ---- PRIMARY buttons only: tiny polish (colour comes from config) ----
           NOTE: number steppers / inputs / selects / sidebar are NOT touched. */
        .stButton>button[kind="primary"],
        .stFormSubmitButton>button[kind="primaryFormSubmit"],
        .stFormSubmitButton>button[kind="primary"]{
            border-radius:10px; font-weight:600;
            transition:transform .15s ease, box-shadow .2s ease, filter .2s ease;
        }
        .stButton>button[kind="primary"]:hover,
        .stFormSubmitButton>button[kind="primaryFormSubmit"]:hover,
        .stFormSubmitButton>button[kind="primary"]:hover{
            transform:translateY(-1px);
            box-shadow:0 6px 14px -4px rgba(2,132,199,.45);
            filter:brightness(1.03);
        }

        /* ---- Tabs: minimal, ice-blue active ---- */
        .stTabs [data-baseweb="tab-list"]{ gap:4px; border-bottom:1px solid var(--kc-border); }
        .stTabs [data-baseweb="tab"]{ font-weight:600; color:var(--kc-muted); }
        .stTabs [aria-selected="true"]{ color:var(--kc-accent-strong); }

        /* Dashboard tables render inside components.html (iframe) with their own
           CSS (see _KC_TABLE_CSS) so click-to-highlight JS can run. */

        /* ---- Mobile ---- */
        @media (max-width:640px){
            .kc-topbar .clock{ display:none; }
            .kc-pill{ margin-left:0; }
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def render_topbar() -> None:
    """Clean, minimalist header (no gradient) with a live Bangkok clock."""
    clock = now_bkk().strftime("%d %b %Y · %H:%M")
    st.markdown(
        f"""
        <div class="kc-topbar">
            <div class="kc-logo">🧊</div>
            <div>
                <div class="t">Koricube Operations Console</div>
                <div class="s">Kori Frost · ระบบจัดการตู้น้ำแข็งอัตโนมัติ</div>
            </div>
            <div class="clock">🕒 Asia/Bangkok · {escape(clock)}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def section_header(icon: str, title: str, subtitle: str) -> None:
    """Reusable minimalist section heading (small icon chip + title)."""
    st.markdown(
        f"""
        <div class="kc-sec">
            <div class="i">{icon}</div>
            <div>
                <div class="h">{escape(title)}</div>
                <div class="d">{escape(subtitle)}</div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_status_card(machine: pd.Series, payment_type: str, is_cash_only: bool) -> None:
    """Compact, single-line status badge: ID · Branch · Merchant · Payment pill."""
    pay_class = "cash" if is_cash_only else "transfer"
    pay_icon = "💵" if is_cash_only else "🔄"
    st.markdown(
        f"""
        <div class="kc-status">
            <div class="it"><span class="k">Machine</span>
                <span class="v">{escape(machine[LOC_MACHINE_ID])}</span></div>
            <div class="sep"></div>
            <div class="it"><span class="k">Branch</span>
                <span class="v">{escape(machine[LOC_BRANCH] or '—')}</span></div>
            <div class="sep"></div>
            <div class="it"><span class="k">Merchant</span>
                <span class="v">{escape(machine[LOC_MERCHANT_NO] or '—')}</span></div>
            <div class="kc-pill {pay_class}">{pay_icon} {escape(payment_type or '—')}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_reconciliation_cue(web_total: Any, cash_collected: Any) -> None:
    """Live colour-coded delta between cash counted and web telemetry total."""
    # None-safe: empty inputs (value=None) are treated as 0.00 for the maths.
    diff = round(_to_float(cash_collected) - _to_float(web_total), 2)
    if abs(diff) < 0.01:
        cls, label, val = "ok", "✅ ตรงกัน (Matched)", "0.00"
    elif diff > 0:
        cls, label, val = "over", "🔺 เงินเกิน (Over)", f"+{diff:,.2f}"
    else:
        cls, label, val = "short", "🔻 เงินขาด (Short)", f"{diff:,.2f}"
    st.markdown(
        f"""
        <div class="kc-rec {cls}">
            <span class="lbl">Reconciliation · Cash − Web</span>
            <span class="val">{label} &nbsp; ฿{val}</span>
        </div>
        """,
        unsafe_allow_html=True,
    )


def machine_label(row: pd.Series) -> str:
    """Dropdown label: 'Branch - Machine_Type (Machine_ID)' (ID keeps it unique)."""
    return f"{row[LOC_BRANCH]} - {row[LOC_MACHINE_TYPE]} ({row[LOC_MACHINE_ID]})"


# ===========================================================================
# UI — FEATURE 1: SALES LOG  (กระทบยอดเงินสด)
# ===========================================================================
def render_sales_log(location_df: pd.DataFrame) -> None:
    section_header("📋", "Sales Log — กระทบยอดเงินสด",
                   "Daily cash reconciliation against web telemetry")

    if location_df.empty:
        st.warning("No machines found in the Location master sheet.")
        return

    # --- Post-submit reset (runs BEFORE any widget is instantiated) ----------
    # NOTE: popping the keys is unreliable — Streamlit's frontend can re-attach
    # the orphaned state, leaving a stale number/date in the field. The robust
    # fix is to EXPLICITLY overwrite each key with its default value here, while
    # we're still ahead of widget instantiation this run.
    if st.session_state.get("sl_do_reset", False):
        st.session_state["sl_remark"] = ""               # Remark -> empty
        st.session_state["sl_cdate"] = now_bkk().date()  # Collection Date -> today
        # Clear every money/period key that may exist (single OR shared mode).
        for _k in ("sl_cash", "sl_web", "sl_pstart", "sl_pend",
                   "sl_web_a", "sl_cash_a", "sl_web_b", "sl_cash_b"):
            if _k in st.session_state:
                st.session_state[_k] = None
        st.session_state["sl_do_reset"] = False          # clear the flag

    # One-shot success flash carried over from the run that just saved + reset.
    _flash = st.session_state.pop("sl_flash", None)
    if _flash:
        st.success(_flash["msg"])
        st.caption(_flash["caption"])

    labels = location_df.apply(machine_label, axis=1).tolist()
    label_to_index = {label: idx for idx, label in enumerate(labels)}

    # Selectbox stays OUTSIDE a form: selecting a machine reruns instantly so the
    # period fields and the live reconciliation cue can react in real time.
    selected_label = st.selectbox("Select Machine (เลือกตู้)", labels)
    machine = location_df.iloc[label_to_index[selected_label]]

    payment_type = machine[LOC_PAYMENT_TYPE]
    is_cash_only = payment_type == PAYMENT_CASH_ONLY
    merchant_no = machine[LOC_MERCHANT_NO]

    # Dynamic mode detection: a machine flagged Machines_Shared == 2 is half of
    # a pair sharing one bank merchant. We auto-locate its partner so the UI can
    # expand to capture BOTH machines and append 2 rows on submit.
    shared_flag = (not is_cash_only) and (
        str(machine[LOC_MACHINES_SHARED]).strip() == "2"
    )
    partner = _find_pair_for(location_df, machine) if shared_flag else None
    is_shared = shared_flag and (partner is not None)

    # Compact, single-line status badge (replaces bulky headings).
    render_status_card(machine, payment_type, is_cash_only)

    # Misconfigured shared machine (flagged 2 but no partner) -> cannot reconcile.
    if shared_flag and partner is None:
        st.error(
            f"⚠️ {machine[LOC_MACHINE_ID]} is flagged Machines_Shared = 2 but no "
            f"partner with Merchant_No {merchant_no or '—'} was found. "
            "Please fix the Location sheet before reconciling."
        )
        return

    if is_shared:
        st.caption(
            f"🔗 Shared merchant {merchant_no or '—'} · pairs with "
            f"{partner[LOC_MACHINE_ID]} ({partner[LOC_BRANCH]}) — 2 rows will be saved."
        )

    # All inputs grouped inside one clean white card.
    with st.container(border=True):
        # Manual, back-datable collection date -> lands in Column A.
        collection_date_obj: date = st.date_input(
            "Collection Date · วันที่เก็บยอด",
            value=now_bkk().date(), key="sl_cdate",
        )

        # Period applies to every transfer machine (single OR shared).
        period_start_obj: Optional[date] = None
        period_end_obj: Optional[date] = None
        if not is_cash_only:
            st.markdown("**🗓️ Reconciliation period · ช่วงรอบบิล**")
            p1, p2 = st.columns(2)
            period_start_obj = p1.date_input("Period Start", value=None, key="sl_pstart")
            period_end_obj = p2.date_input("Period End", value=None, key="sl_pend")

        st.markdown("**💰 Amounts · ยอดเงิน**")
        # Empty-by-default inputs (value=None + "0.00" placeholder) so the
        # operator can type immediately without clearing a pre-filled 0.00.
        web_total = cash_collected = None         # single-machine values
        web_a = cash_a = web_b = cash_b = None    # shared-pair values
        machine_a = machine_b = None

        if is_cash_only:
            # Cash-only: no telemetry, no Web Total.
            cash_collected = st.number_input(
                "Cash Collected · เงินสดที่เก็บได้ (฿)",
                value=None, min_value=0.0, step=1.0, format="%.2f",
                placeholder="0.00", key="sl_cash",
            )
        elif is_shared:
            # Shared: capture BOTH machines (A = selected, B = partner).
            machine_a, machine_b = machine, partner
            st.markdown(
                f"**🅰️ {escape(str(machine_a[LOC_MACHINE_ID]))}** · "
                f"{escape(str(machine_a[LOC_BRANCH]))}"
            )
            a1, a2 = st.columns(2)
            web_a = a1.number_input("Web Total A (฿)", value=None, min_value=0.0,
                                    step=1.0, format="%.2f", placeholder="0.00",
                                    key="sl_web_a")
            cash_a = a2.number_input("Cash Collected A (฿)", value=None, min_value=0.0,
                                     step=1.0, format="%.2f", placeholder="0.00",
                                     key="sl_cash_a")
            st.markdown(
                f"**🅱️ {escape(str(machine_b[LOC_MACHINE_ID]))}** · "
                f"{escape(str(machine_b[LOC_BRANCH]))}"
            )
            b1, b2 = st.columns(2)
            web_b = b1.number_input("Web Total B (฿)", value=None, min_value=0.0,
                                    step=1.0, format="%.2f", placeholder="0.00",
                                    key="sl_web_b")
            cash_b = b2.number_input("Cash Collected B (฿)", value=None, min_value=0.0,
                                     step=1.0, format="%.2f", placeholder="0.00",
                                     key="sl_cash_b")
        else:
            # Single transfer machine.
            m1, m2 = st.columns(2)
            web_total = m1.number_input(
                "Web Total · ยอดเว็บ (฿)",
                value=None, min_value=0.0, step=1.0, format="%.2f",
                placeholder="0.00", key="sl_web",
            )
            cash_collected = m2.number_input(
                "Cash Collected · เงินสดที่เก็บได้ (฿)",
                value=None, min_value=0.0, step=1.0, format="%.2f",
                placeholder="0.00", key="sl_cash",
            )
            render_reconciliation_cue(web_total, cash_collected)

        remark = st.text_area(
            "Remark · หมายเหตุ", placeholder="Optional notes…", key="sl_remark"
        )

        # --- Auto-aggregate the bank transfer (X, Y) over the PERIOD ----------
        # X = Σ Trans_Amount (gross), Y = Σ Commission, summed across the merchant's
        # settlements in [Period_Start, Period_End]. No manual bank entry.
        collection_date = date_to_iso(collection_date_obj) or today_iso()
        period_start = "" if is_cash_only else date_to_iso(period_start_obj)
        period_end = "" if is_cash_only else date_to_iso(period_end_obj)

        bank: Optional[tuple] = None
        result: Optional[SharedReconResult] = None
        if not is_cash_only:
            if not period_start or not period_end:
                st.info(
                    "Select Period Start and Period End to load the bank transfer "
                    "for this period."
                )
            else:
                bank = safe_aggregate_bank_transfers(merchant_no, period_start, period_end)
                if bank is None:
                    st.warning(
                        f"⏳ Bank data for period {period_start} → {period_end} · "
                        f"Merchant {merchant_no or '—'} is missing in Raw_Email — "
                        "ยังไม่มีข้อมูลธนาคารสำหรับช่วงนี้. Submit is disabled."
                    )
                else:
                    x_val, y_val = bank
                    if is_shared:
                        result = compute_shared_reconciliation(
                            web_a, cash_a, web_b, cash_b, x_val, y_val, shares=2
                        )
                        render_shared_preview(machine_a, machine_b, result)
                    else:
                        bk1, bk2 = st.columns(2)
                        bk1.metric("Actual Transfer · X = Σ Trans_Amount", f"฿{x_val:,.2f}")
                        bk2.metric("Bank Fee · Y = Σ Commission", f"฿{y_val:,.2f}")

        # Transfer machines need the matched bank data before saving.
        submit_disabled = (not is_cash_only) and (bank is None)
        submitted = st.button(
            "Submit Sales Log", type="primary", use_container_width=True,
            key="sl_submit", disabled=submit_disabled,
        )

    if not submitted:
        return

    # Validate the period for transfer machines (single + shared).
    if not is_cash_only:
        if not period_start or not period_end:
            st.error("Please provide both Period Start and Period End.")
            return
        if period_start > period_end:
            st.error("Period Start must not be after Period End.")
            return
        if bank is None:
            st.error("Bank data is not available for this period.")
            return

    # --- Build the row(s): 2 for a shared pair, 1 otherwise (uniform A→P) -----
    if is_shared:
        rows = build_shared_reconciliation_rows(
            collection_date=collection_date, period_start=period_start,
            period_end=period_end, remark=remark.strip(),
            machine_a=machine_a, machine_b=machine_b,
            web_a=web_a, cash_a=cash_a, web_b=web_b, cash_b=cash_b, result=result,
        )
    else:
        rows = [build_single_sales_row(
            collection_date=collection_date, machine=machine,
            web_total=web_total, cash_collected=cash_collected,
            period_start=period_start, period_end=period_end,
            remark=remark.strip(), is_cash_only=is_cash_only, bank=bank,
        )]

    # Append every row (append_row only -> row 1 ArrayFormula untouched).
    try:
        for row in rows:
            append_row_safe(WS_SALES_LOG, row)
    except Exception as exc:  # noqa: BLE001 - present any backend error nicely
        st.error(f"Failed to write to Sales_Log: {exc}")
        return

    # One-shot success flash + reset, then rerun so the form comes back cleared.
    if is_shared:
        st.session_state["sl_flash"] = {
            "msg": (f"✅ Saved 2 rows (shared) · Merchant {merchant_no}: "
                    f"{machine_a[LOC_MACHINE_ID]} + {machine_b[LOC_MACHINE_ID]}."),
            "caption": (f"Net A ฿{result.net_actual_a:,.2f} · "
                        f"Net B ฿{result.net_actual_b:,.2f} · {collection_date}"),
        }
    else:
        st.session_state["sl_flash"] = {
            "msg": f"✅ Sales log saved for {machine[LOC_MACHINE_ID]} "
                   f"({machine[LOC_BRANCH]}).",
            "caption": f"Collection date stamped: {collection_date}",
        }
    st.session_state["sl_do_reset"] = True
    st.rerun()


# ===========================================================================
# SALES LOG HELPERS — pair lookup · single-row builder · shared preview
# ===========================================================================
def _find_pair_for(location_df: pd.DataFrame, machine: pd.Series) -> Optional[pd.Series]:
    """
    Given a shared machine (Machines_Shared == 2), return its PARTNER — the
    other machine sharing the same Merchant_No — or ``None`` if not found.
    """
    merchant = _digits_only(machine[LOC_MERCHANT_NO])
    if not merchant:
        return None
    same = location_df[
        (location_df[LOC_MERCHANT_NO].apply(_digits_only) == merchant)
        & (location_df[LOC_MACHINE_ID] != machine[LOC_MACHINE_ID])
        & (location_df[LOC_MACHINES_SHARED].astype(str).str.strip() == "2")
    ]
    return same.iloc[0] if not same.empty else None


def build_single_sales_row(
    *, collection_date: str, machine: pd.Series, web_total: Any, cash_collected: Any,
    period_start: str, period_end: str, remark: str,
    is_cash_only: bool, bank: Optional[tuple],
) -> List[Any]:
    """
    Build ONE full 16-column Sales_Log row (A→P) for a single machine.

    Single transfer machine (per spec): Diff_Adjustment = 0, Shared_Bank_Fee = Y,
    Net_Actual = X, Expected_Transfer = Web − Cash.
    Cash-only machine: no bank transfer -> L–P blank, Status = cash.
    """
    if is_cash_only:
        web_value: Any = ""
        expected = shared_fee = net_actual = diff_adj = ""
        status = RECON_STATUS_CASH
    else:
        x_val, y_val = bank  # guaranteed present (submit was enabled)
        web_value = _to_float(web_total)
        expected = round(_to_float(web_total) - _to_float(cash_collected), 2)
        shared_fee = round(_to_float(y_val), 2)   # M  Shared_Bank_Fee = Y
        net_actual = round(_to_float(x_val), 2)   # N  Net_Actual = X
        diff_adj = 0                              # P  Diff_Adjustment = 0
        status = RECON_STATUS_DONE

    return [
        collection_date,               # A  Collection_Date
        machine[LOC_MACHINE_ID],       # B  Machine_ID
        machine[LOC_BRANCH],           # C  Branch_Name
        machine[LOC_MERCHANT_NO],      # D  Merchant_No
        machine[LOC_PAYMENT_TYPE],     # E  Payment_Type
        machine[LOC_MACHINES_SHARED],  # F  Machines_Shared
        web_value,                     # G  Web_Total ("" for cash-only)
        _to_float(cash_collected),     # H  Cash_Collected
        period_start,                  # I  Period_Start
        period_end,                    # J  Period_End
        remark,                        # K  Remark
        expected,                      # L  Expected_Transfer
        shared_fee,                    # M  Shared_Bank_Fee (= Y)
        net_actual,                    # N  Net_Actual (= X)
        status,                        # O  Status
        diff_adj,                      # P  Diff_Adjustment
    ]


def render_shared_preview(machine_a: pd.Series, machine_b: pd.Series,
                          r: SharedReconResult) -> None:
    """Transparent, live breakdown of the shared reconciliation maths."""
    st.markdown("**🔎 Reconciliation preview · ตรวจสอบการคำนวณ**")

    s1, s2, s3, s4 = st.columns(4)
    s1.metric("Total Expected", f"฿{r.total_expected:,.2f}")
    s2.metric("Actual Transfer (X)", f"฿{r.actual_transfer:,.2f}")
    s3.metric("Diff · X − Exp", f"฿{r.diff:,.2f}")
    s4.metric("Total Fee (Y)", f"฿{r.total_fee:,.2f}")

    s5, s6 = st.columns(2)
    s5.metric("Diff ÷ 2 · Diff_Adjustment", f"฿{r.diff_adjustment:,.2f}")
    s6.metric("Fee ÷ 2 · Shared_Bank_Fee", f"฿{r.shared_bank_fee:,.2f}")

    n1, n2 = st.columns(2)
    n1.metric(
        f"🅰️ Net Actual · {machine_a[LOC_MACHINE_ID]}",
        f"฿{r.net_actual_a:,.2f}",
        delta=f"Expected ฿{r.expected_a:,.2f}", delta_color="off",
    )
    n2.metric(
        f"🅱️ Net Actual · {machine_b[LOC_MACHINE_ID]}",
        f"฿{r.net_actual_b:,.2f}",
        delta=f"Expected ฿{r.expected_b:,.2f}", delta_color="off",
    )


# ===========================================================================
# UI — FEATURE 3: MAINTENANCE & EXPENSE LEDGER  (แจ้งซ่อม / บันทึกรายจ่าย)
# ===========================================================================
def _is_expense_code(code: str) -> bool:
    """True for variable-expense codes (utility bills / misc), not technical repairs."""
    return code in EXPENSE_CODES


def build_maintenance_row(
    *, report_date: str, machine: pd.Series, code: str, desc: str,
    cost: Any, is_expense: bool
) -> List[Any]:
    """
    Build ONE Maintenance row using the 8-column schema (MAINTENANCE_COLUMNS):
    [Report_Date, Machine_ID, Branch_Name, Error_Code, Issue_Desc,
     Repair_Cost, Resolved_Date, Status].

    ``report_date`` is the user-chosen date (YYYY-MM-DD) — the bill date/period
    for expenses, or the report date for repairs. Branch_Name is stamped from
    the CURRENT Location data. Repairs default to pending (รอซ่อม) with a blank
    Resolved_Date; expense entries are settled on the bill date (เคลียร์แล้ว) so
    Resolved_Date mirrors report_date.
    """
    if is_expense:
        status, resolved_date = STATUS_EXPENSE_PAID, report_date
    else:
        status, resolved_date = STATUS_PENDING, ""
    return [
        report_date,               # A  Report_Date (bill date for expenses)
        machine[LOC_MACHINE_ID],   # B  Machine_ID
        machine[LOC_BRANCH],       # C  Branch_Name (stamped from Location)
        code,                      # D  Error_Code (technical code OR expense type)
        desc.strip(),              # E  Issue_Desc / Description
        _to_float(cost),           # F  Repair_Cost / amount (None -> 0.0)
        resolved_date,             # G  Resolved_Date
        status,                    # H  Status
    ]


def render_maintenance(location_df: pd.DataFrame) -> None:
    section_header("🧾", "Maintenance & Expenses — แจ้งซ่อม / บันทึกรายจ่าย",
                   "Repairs + variable expenses (water, electricity, misc)")

    if location_df.empty:
        st.warning("No machines found in the Location master sheet.")
        return

    labels = location_df.apply(machine_label, axis=1).tolist()
    label_to_index = {label: idx for idx, label in enumerate(labels)}

    with st.form("maintenance_form", clear_on_submit=True):
        top1, top2 = st.columns([2, 1])
        selected_label = top1.selectbox("Select Machine (เลือกตู้)", labels)
        code = top2.selectbox(
            "Error Code / Expense Type (ประเภทรายจ่าย)", MAINTENANCE_CODES,
            format_func=lambda c: MAINTENANCE_CODE_LABELS.get(c, c),
        )
        # Back-datable date -> lands in Column A. For utility bills this is the
        # bill date/month; for repairs it's the report date.
        report_date_obj = st.date_input(
            "Date / วันที่ (รอบบิลค่าน้ำ–ค่าไฟ)", value=now_bkk().date(),
        )
        desc = st.text_area("Description / บันทึกเพิ่มเติม")
        cost = st.number_input(
            "Cost / ยอดเงิน (฿)", value=None, min_value=0.0, step=1.0,
            format="%.2f", placeholder="0.00",
        )
        submitted = st.form_submit_button(
            "Submit Entry", type="primary", use_container_width=True
        )

    if not submitted:
        return

    is_expense = _is_expense_code(code)
    report_date = date_to_iso(report_date_obj) or today_iso()
    machine = location_df.iloc[label_to_index[selected_label]]
    payload = build_maintenance_row(
        report_date=report_date, machine=machine, code=code, desc=desc,
        cost=cost, is_expense=is_expense,
    )

    try:
        append_row_safe(WS_MAINTENANCE, payload)
    except Exception as exc:  # noqa: BLE001
        st.error(f"Failed to write to Maintenance: {exc}")
        return

    kind = "Expense" if is_expense else "Repair ticket"
    st.success(
        f"✅ {kind} saved for {machine[LOC_MACHINE_ID]} · {machine[LOC_BRANCH]} · "
        f"{code} (status: {payload[-1]})."
    )


# ===========================================================================
# DASHBOARD — Hybrid data prep (legacy + current) → interactive analytics
# ===========================================================================
# Master grain is (Machine_ID, Month). Revenue comes from Sales_Log, the
# variable bills from Maintenance, fixed Rent is mapped once-per-machine-month
# from Location, and the legacy Final_Database supplies pre-calculated history.
DASH_NUMERIC = [
    "Cash_Revenue", "Transfer_Revenue",
    "Rent_Cost", "Water_Cost", "Electric_Cost", "Maintenance_Cost",
]
DASH_COLUMNS = [
    "Month", "Month_Label", "Machine_ID", "Branch_Name", "Machine_Type",
    "Cash_Revenue", "Transfer_Revenue", "Total_Revenue",
    "Rent_Cost", "Water_Cost", "Electric_Cost", "Maintenance_Cost",
    "Total_Expenses", "Net_Profit", "Source",
]


def _fetch_records_df(ws_name: str) -> pd.DataFrame:
    """
    Read a worksheet into a DataFrame; empty on any failure (missing sheet /
    duplicate headers). UNFORMATTED_VALUE returns real numbers (so a Repair_Cost
    cell mis-formatted as a date still reads as its number) and dates as serials,
    which ``_to_month_start`` handles explicitly.
    """
    try:
        ws = get_worksheet(ws_name)
        try:
            records = ws.get_all_records(value_render_option="UNFORMATTED_VALUE")
        except TypeError:  # very old gspread without the kwarg
            records = ws.get_all_records()
        return pd.DataFrame(records)
    except Exception:  # noqa: BLE001
        return pd.DataFrame()


def _series_num(df: pd.DataFrame, *names: str, default: float = 0.0) -> pd.Series:
    """First matching column coerced to float; a default-filled series if absent."""
    for name in names:
        if name in df.columns:
            return df[name].map(_to_float)
    return pd.Series(default, index=df.index, dtype=float)


def _series_str(df: pd.DataFrame, *names: str, default: str = "") -> pd.Series:
    """First matching column as stripped strings; a default-filled series if absent."""
    for name in names:
        if name in df.columns:
            return df[name].astype(str).str.strip()
    return pd.Series(default, index=df.index, dtype=object)


def _series_raw(df: pd.DataFrame, *names: str) -> pd.Series:
    """
    First matching column UNCHANGED (no string coercion); all-None if absent.

    Used for date columns so ``_to_month_start`` still sees numeric Google-Sheets
    serials (UNFORMATTED_VALUE) instead of pre-stringified text.
    """
    for name in names:
        if name in df.columns:
            return df[name]
    return pd.Series([None] * len(df), index=df.index, dtype=object)


def _to_datetime(values: pd.Series) -> pd.Series:
    """
    Parse a date-ish column to day-level Timestamps (NaT for blanks/garbage).

    Handles three shapes robustly: ISO/locale strings (incl. Buddhist Era),
    Google-Sheets serial numbers (UNFORMATTED dates), and blanks.
    """
    epoch = pd.Timestamp("1899-12-30")  # Google Sheets day 0

    def parse(v: Any):
        if isinstance(v, bool):
            return pd.NaT
        if isinstance(v, (int, float)):
            return epoch + pd.to_timedelta(int(v), unit="D") if 1 <= v <= 80000 else pd.NaT
        return pd.to_datetime(normalize_date_string(v), errors="coerce")

    return pd.to_datetime(values.map(parse), errors="coerce")


def _to_month_start(values: pd.Series) -> pd.Series:
    """Day-level parse (see ``_to_datetime``) floored to the first of the month."""
    return _to_datetime(values).dt.to_period("M").dt.to_timestamp()


def _machine_branch_map(*frames: pd.DataFrame) -> dict:
    """Machine_ID -> Branch_Name (first non-empty wins across the given frames)."""
    mapping: dict = {}
    for df in frames:
        if df is None or df.empty:
            continue
        for mid, branch in zip(_series_str(df, "Machine_ID"),
                               _series_str(df, "Branch_Name", LOC_BRANCH)):
            if mid and branch and mid not in mapping:
                mapping[mid] = branch
    return mapping


def _machine_type_map(df_loc: pd.DataFrame) -> dict:
    """Machine_ID -> Machine_Type (e.g. Water / Ice) from the Location sheet."""
    if df_loc.empty:
        return {}
    return {mid: typ for mid, typ in
            zip(_series_str(df_loc, LOC_MACHINE_ID),
                _series_str(df_loc, LOC_MACHINE_TYPE)) if mid and typ}


def _location_rent_df(df_loc: pd.DataFrame) -> pd.DataFrame:
    """One fixed monthly rent per Machine_ID (0 if the column is absent)."""
    if df_loc.empty:
        return pd.DataFrame(columns=["Machine_ID", "Rent_Cost"])
    return pd.DataFrame({
        "Machine_ID": _series_str(df_loc, LOC_MACHINE_ID),
        "Rent_Cost": _series_num(df_loc, LOC_RENT_COST, "Rent"),
    }).drop_duplicates("Machine_ID")


def _prep_sales_revenue(df_sl: pd.DataFrame) -> pd.DataFrame:
    """Sales_Log -> per-(Machine_ID, Month) Cash_Revenue + Transfer_Revenue."""
    cols = ["Machine_ID", "Month", "Branch_Name", "Cash_Revenue", "Transfer_Revenue"]
    if df_sl.empty:
        return pd.DataFrame(columns=cols)
    raw = pd.DataFrame({
        "Machine_ID": _series_str(df_sl, "Machine_ID"),
        "Branch_Name": _series_str(df_sl, "Branch_Name"),
        "Month": _to_month_start(_series_raw(df_sl, "Collection_Date")),
        "Cash_Revenue": _series_num(df_sl, "Cash_Collected"),
        "Transfer_Revenue": _series_num(df_sl, "Net_Actual"),
    }).dropna(subset=["Month"])
    raw = raw[raw["Machine_ID"] != ""]
    return raw.groupby(["Machine_ID", "Month"], as_index=False).agg(
        Branch_Name=("Branch_Name", "first"),
        Cash_Revenue=("Cash_Revenue", "sum"),
        Transfer_Revenue=("Transfer_Revenue", "sum"),
    )


def _prep_maintenance_costs(df_mt: pd.DataFrame) -> pd.DataFrame:
    """Maintenance -> per-(Machine_ID, Month) Water / Electric / Maintenance costs."""
    cols = ["Machine_ID", "Month", "Water_Cost", "Electric_Cost", "Maintenance_Cost"]
    if df_mt.empty:
        return pd.DataFrame(columns=cols)
    code = _series_str(df_mt, "Error_Code")
    cost = _series_num(df_mt, "Repair_Cost")
    raw = pd.DataFrame({
        "Machine_ID": _series_str(df_mt, "Machine_ID"),
        "Month": _to_month_start(_series_raw(df_mt, "Report_Date")),
        "Water_Cost": cost.where(code == "BILL-WATER", 0.0),
        "Electric_Cost": cost.where(code == "BILL-ELEC", 0.0),
        "Maintenance_Cost": cost.where(~code.isin(["BILL-WATER", "BILL-ELEC"]), 0.0),
    }).dropna(subset=["Month"])
    raw = raw[raw["Machine_ID"] != ""]
    return raw.groupby(["Machine_ID", "Month"], as_index=False)[
        ["Water_Cost", "Electric_Cost", "Maintenance_Cost"]
    ].sum()


def _assemble_current(df_sl: pd.DataFrame, df_mt: pd.DataFrame,
                      rent_df: pd.DataFrame, branch_map: dict) -> pd.DataFrame:
    """Merge current revenue + bills, stamp rent once per machine-month."""
    rev = _prep_sales_revenue(df_sl)
    exp = _prep_maintenance_costs(df_mt)
    if rev.empty and exp.empty:
        return pd.DataFrame()

    cur = pd.merge(rev, exp, on=["Machine_ID", "Month"], how="outer")
    for c in ["Cash_Revenue", "Transfer_Revenue",
              "Water_Cost", "Electric_Cost", "Maintenance_Cost"]:
        cur[c] = pd.to_numeric(cur[c], errors="coerce").fillna(0.0) if c in cur else 0.0
    if "Branch_Name" not in cur:
        cur["Branch_Name"] = ""
    blank = cur["Branch_Name"].isna() | (cur["Branch_Name"].astype(str).str.strip() == "")
    cur.loc[blank, "Branch_Name"] = cur.loc[blank, "Machine_ID"].map(branch_map)

    cur = cur.merge(rent_df, on="Machine_ID", how="left")
    cur["Rent_Cost"] = pd.to_numeric(cur.get("Rent_Cost"), errors="coerce").fillna(0.0)
    cur["Source"] = "current"
    return cur


def _prep_legacy(df_final: pd.DataFrame) -> pd.DataFrame:
    """Final_Database -> master columns (defensive candidate-name lookups)."""
    if df_final.empty:
        return pd.DataFrame()
    out = pd.DataFrame({
        "Machine_ID": _series_str(df_final, "Machine_ID"),
        "Branch_Name": _series_str(df_final, "Branch_Name", LOC_BRANCH),
        "Month": _to_month_start(
            _series_raw(df_final, "Date", "Month", "Collection_Date", "Report_Date")
        ),
        "Cash_Revenue": _series_num(df_final, "Cash_Revenue", "Cash", "Cash_Collected"),
        "Transfer_Revenue": _series_num(df_final, "Transfer_Revenue", "Net_Actual", "Transfer"),
        "Rent_Cost": _series_num(df_final, "Rent_Cost", "Rent"),
        "Water_Cost": _series_num(df_final, "Water_Cost", "Water"),
        "Electric_Cost": _series_num(df_final, "Electric_Cost", "Electric", "Elec"),
        "Maintenance_Cost": _series_num(df_final, "Maintenance_Cost", "Maintenance", "Repair_Cost"),
    }).dropna(subset=["Month"])
    out["Source"] = "legacy"
    return out


def _dedupe_machine_month(master: pd.DataFrame) -> pd.DataFrame:
    """If a (Machine_ID, Month) exists in both sources, keep the current one."""
    if master.empty:
        return master
    master = master.copy()
    master["_pri"] = (master["Source"] == "current").astype(int)
    return (master.sort_values("_pri", ascending=False)
                  .drop_duplicates(subset=["Machine_ID", "Month"], keep="first")
                  .drop(columns="_pri"))


@st.cache_data(ttl=120, show_spinner="Building dashboard…")
def load_and_prep_dashboard_data() -> pd.DataFrame:
    """Single tidy master DataFrame (legacy ⊕ current) for the dashboard."""
    df_sl = _fetch_records_df(WS_SALES_LOG)
    df_mt = _fetch_records_df(WS_MAINTENANCE)
    df_final = _fetch_records_df(WS_FINAL_DB)
    df_loc = _fetch_records_df(WS_LOCATION)

    branch_map = _machine_branch_map(df_sl, df_mt, df_loc)
    current = _assemble_current(df_sl, df_mt, _location_rent_df(df_loc), branch_map)
    legacy = _prep_legacy(df_final)

    master = pd.concat([legacy, current], ignore_index=True, sort=False)
    if master.empty:
        return pd.DataFrame(columns=DASH_COLUMNS)

    for c in DASH_NUMERIC:
        master[c] = pd.to_numeric(master[c], errors="coerce").fillna(0.0) if c in master else 0.0
    master["Machine_ID"] = _series_str(master, "Machine_ID")
    master["Branch_Name"] = (_series_str(master, "Branch_Name")
                             .replace("", "ไม่ระบุสาขา"))
    # Map machine type (Water / Ice) from Location for the drill-down slicer.
    type_map = _machine_type_map(df_loc)
    master["Machine_Type"] = (master["Machine_ID"].map(type_map)
                              .fillna("").replace("", "ไม่ระบุ"))

    master = _dedupe_machine_month(master)
    master["Total_Revenue"] = master["Cash_Revenue"] + master["Transfer_Revenue"]
    master["Total_Expenses"] = master[
        ["Rent_Cost", "Water_Cost", "Electric_Cost", "Maintenance_Cost"]
    ].sum(axis=1)
    master["Net_Profit"] = master["Total_Revenue"] - master["Total_Expenses"]
    master["Month_Label"] = master["Month"].dt.strftime("%Y-%m")
    return master[DASH_COLUMNS].sort_values(["Branch_Name", "Month"]).reset_index(drop=True)


# ===========================================================================
# BANK TRANSFERS — Raw_Email settlements NOT yet reconciled in Sales_Log
# ===========================================================================
PENDING_VALUE_COLS = ["Trans_Amount", "Commission", "VAT", "Net_Transfer"]


def _merchant_branch_map(location_df: pd.DataFrame) -> dict:
    """
    Merchant_No (digits) -> readable label of its branch(es). A shared merchant
    that maps to two branches is joined with ' + ' (e.g. 'ป.พัน7 + เจ็ดยอด3'); the
    raw Merchant_No is appended for traceability.
    """
    if location_df is None or location_df.empty:
        return {}
    by_merchant: dict = {}
    for raw_m, branch in zip(_series_str(location_df, LOC_MERCHANT_NO),
                             _series_str(location_df, LOC_BRANCH, "Branch_Name")):
        merchant = _digits_only(raw_m)
        if not merchant:
            continue
        names = by_merchant.setdefault(merchant, [])
        if branch and branch not in names:
            names.append(branch)
    return {m: " + ".join(sorted(names)) for m, names in by_merchant.items()}


def _sales_log_periods(df_sl: pd.DataFrame) -> dict:
    """
    {merchant_digits: [(start_ts, end_ts), ...]} for every Sales_Log entry that
    carries a merchant + both period dates — i.e. the windows already reconciled.
    """
    periods: dict = {}
    if df_sl is None or df_sl.empty:
        return periods
    starts = _to_datetime(_series_raw(df_sl, "Period_Start"))
    ends = _to_datetime(_series_raw(df_sl, "Period_End"))
    merchants = _series_str(df_sl, "Merchant_No")
    for raw_m, start, end in zip(merchants, starts, ends):
        merchant = _digits_only(raw_m)
        if merchant and not pd.isna(start) and not pd.isna(end):
            periods.setdefault(merchant, []).append((start, end))
    return periods


def _is_reconciled(merchant: str, when, periods: dict) -> bool:
    """True if ``when`` (a Timestamp) falls within any reconciled period."""
    if pd.isna(when):
        return False
    return any(start <= when <= end for start, end in periods.get(merchant, []))


@st.cache_data(ttl=120, show_spinner="Loading pending bank transfers…")
def load_pending_transfers() -> pd.DataFrame:
    """
    Raw_Email rows NOT yet reconciled in Sales_Log, aggregated per
    (Merchant, Month). A row is reconciled when its Transfer_Date sits inside a
    Sales_Log [Period_Start, Period_End] for the same merchant.

    Columns: Merchant_No, Merchant_Label, Month, Month_Label + summed
    Trans_Amount / Commission / VAT / Net_Transfer. Empty (well-formed) when
    nothing is pending or the sheets are unreadable.
    """
    cols = ["Merchant_No", "Merchant_Label", "Month", "Month_Label", "Last_Update",
            *PENDING_VALUE_COLS]
    try:
        raw = fetch_raw_email_data()
    except Exception:  # noqa: BLE001
        return pd.DataFrame(columns=cols)
    if raw is None or raw.empty:
        return pd.DataFrame(columns=cols)

    df_sl = _fetch_records_df(WS_SALES_LOG)
    df_loc = _fetch_records_df(WS_LOCATION)
    periods = _sales_log_periods(df_sl)
    label_map = _merchant_branch_map(df_loc)

    tidy = pd.DataFrame({
        "Merchant_No": raw["Merchant_No"].apply(_digits_only),
        "_date": _to_datetime(_series_raw(raw, "Transfer_Date")),
        "Month": _to_month_start(_series_raw(raw, "Transfer_Date")),
        "Trans_Amount": _series_num(raw, "Trans_Amount"),
        "Commission": _series_num(raw, "Commission"),
        "VAT": _series_num(raw, "VAT"),
        "Net_Transfer": _series_num(raw, "Net_Transfer"),
    }).dropna(subset=["Month"])
    tidy = tidy[tidy["Merchant_No"] != ""]
    if tidy.empty:
        return pd.DataFrame(columns=cols)

    pending_mask = ~tidy.apply(
        lambda r: _is_reconciled(r["Merchant_No"], r["_date"], periods), axis=1)
    pending = tidy[pending_mask]
    if pending.empty:
        return pd.DataFrame(columns=cols)

    agg = pending.groupby(["Merchant_No", "Month"], as_index=False).agg(
        Trans_Amount=("Trans_Amount", "sum"),
        Commission=("Commission", "sum"),
        VAT=("VAT", "sum"),
        Net_Transfer=("Net_Transfer", "sum"),
        _last=("_date", "max"),  # most recent Transfer_Date in this bucket
    )
    agg["Merchant_Label"] = agg["Merchant_No"].map(label_map).fillna("")
    agg["Month_Label"] = agg["Month"].dt.strftime("%Y-%m")
    agg["Last_Update"] = agg["_last"].dt.strftime("%Y-%m-%d")
    return agg[cols].sort_values(["Merchant_Label", "Month"]).reset_index(drop=True)


# --- Dashboard UI pieces ---------------------------------------------------
def _dashboard_kpis(view: pd.DataFrame) -> None:
    revenue = float(view["Total_Revenue"].sum())
    expenses = float(view["Total_Expenses"].sum())
    profit = float(view["Net_Profit"].sum())
    margin = (profit / revenue * 100) if revenue else 0.0
    c1, c2, c3 = st.columns(3)
    c1.metric("Total Revenue · รายรับรวม", _baht(revenue))
    c2.metric("Total Expenses · รายจ่ายรวม", _baht(expenses))
    c3.metric("Net Profit · กำไรสุทธิ", _baht(profit), delta=f"{margin:,.1f}% margin")


def _profit_css(value: float) -> str:
    """Green text for positive net profit, red for negative, plain for zero."""
    if value > 0:
        return "color: #16a34a; font-weight: 600"   # green
    if value < 0:
        return "color: #dc2626; font-weight: 600"   # red
    return ""


def _negate_expense(v: float) -> float:
    """Show a cost as a negative deduction (exact zero stays 0.0, not -0.0)."""
    return -abs(v) if abs(v) > 1e-9 else 0.0


def _baht(value: float) -> str:
    """Money as a comma-grouped string with a trailing 'บาท' (e.g. 15,071.12 บาท)."""
    return f"{value:,.2f} บาท"


def _comma(value: float) -> str:
    """Comma-grouped 2-dp number, NO currency suffix (unit lives in the header)."""
    return f"{value:,.2f}"


# Self-contained CSS for the iframe-rendered tables (literal hex, no :root vars
# since the iframe doesn't inherit the page theme).
_KC_TABLE_CSS = """
  @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=Prompt:wght@400;500;600&display=swap');
  *{ box-sizing:border-box; }
  body{ margin:0; background:transparent; color:#0F172A;
        font-family:'Inter','Prompt',system-ui,-apple-system,sans-serif;
        font-variant-numeric:tabular-nums;
        -webkit-font-smoothing:antialiased; text-rendering:optimizeLegibility; }
  .kc-table{ max-height:460px; overflow:auto; -webkit-overflow-scrolling:touch;
             border:1px solid #E2E8F0; border-radius:14px; background:#fff; }
  .kc-table table{ border-collapse:collapse; width:100%; font-size:.86rem; }
  .kc-table thead th{ background:#f7f8fc; color:#64748B; font-weight:600;
       padding:9px 14px; text-align:right; white-space:nowrap;
       border-bottom:1px solid #E2E8F0; position:sticky; top:0; z-index:2; }
  .kc-table tbody td{ padding:8px 14px; text-align:right; white-space:nowrap;
       border-bottom:1px solid #E2E8F0; color:#0F172A; }
  .kc-table th:first-child, .kc-table td:first-child{ text-align:left; }
  .kc-table tbody tr:last-child td{ border-bottom:none; }
  /* Click/tap to highlight a row — soft slate tones, easy on the eyes. */
  .kc-table tbody tr{ cursor:pointer; }
  .kc-table tbody tr:hover td{ background:#f1f5f9; }
  .kc-table tbody tr.sel td{ background:#e2e8f0; }
  /* Freeze first two columns (สาขา + เดือน) on horizontal scroll. */
  .kc-freeze2 th:nth-child(1), .kc-freeze2 td:nth-child(1){
       position:sticky; left:0; width:118px; min-width:118px; max-width:118px;
       overflow:hidden; text-overflow:ellipsis; }
  .kc-freeze2 th:nth-child(2), .kc-freeze2 td:nth-child(2){
       position:sticky; left:118px; width:82px; min-width:82px; max-width:82px;
       box-shadow:2px 0 0 0 #E2E8F0; }
  .kc-freeze2 tbody td:nth-child(1), .kc-freeze2 tbody td:nth-child(2){
       background:#fff; z-index:1; }
  .kc-freeze2 thead th:nth-child(1), .kc-freeze2 thead th:nth-child(2){ z-index:3; }
  /* Freeze only the first column (e.g. a wide merchant label). */
  .kc-freeze1 th:nth-child(1), .kc-freeze1 td:nth-child(1){
       position:sticky; left:0; width:170px; min-width:170px; max-width:170px;
       overflow:hidden; text-overflow:ellipsis; box-shadow:2px 0 0 0 #E2E8F0; }
  .kc-freeze1 tbody td:nth-child(1){ background:#fff; z-index:1; }
  .kc-freeze1 thead th:nth-child(1){ z-index:3; }
"""


def _render_styled_table(styler, freeze: int = 0) -> None:
    """
    Render a pandas Styler as a real HTML table inside ``components.html`` (an
    iframe), so click/tap-to-highlight JS actually runs.

    width:100% fits & fills the desktop container (no horizontal scroll); on a
    narrow phone the no-wrap cells scroll instead of truncating numbers. Keeps
    the Styler's commas / 'บาท' / green-red colours. ``freeze=2`` pins the first
    two columns so the branch/month stay visible while the rest scrolls.
    """
    table_html = styler.hide(axis="index").to_html()
    css_class = "kc-table" + (f" kc-freeze{freeze}" if freeze else "")
    n_rows = int(styler.data.shape[0])
    height = min(48 + n_rows * 39, 472)  # fit small tables; cap + inner-scroll big ones
    doc = f"""<style>{_KC_TABLE_CSS}</style>
<div class="{css_class}">{table_html}</div>
<script>
(function(){{
  document.querySelectorAll('.kc-table tbody tr').forEach(function(r){{
    r.addEventListener('click', function(){{ r.classList.toggle('sel'); }});
  }});
}})();
</script>"""
    components.html(doc, height=height, scrolling=False)


def _dashboard_trend_chart(view: pd.DataFrame) -> None:
    monthly = (view.groupby(["Month", "Month_Label"], as_index=False)
                   .agg(Total_Revenue=("Total_Revenue", "sum"),
                        Net_Profit=("Net_Profit", "sum"))
                   .sort_values("Month"))
    if monthly.empty:
        return

    # Discrete, evenly-spaced MONTHLY axis (one label per month, chronological)
    # instead of a temporal axis that auto-ticks by week and looks cluttered.
    monthly["Month_Disp"] = monthly["Month"].dt.strftime("%b %Y")
    x = alt.X("Month_Disp:O", title="เดือน", sort=monthly["Month_Disp"].tolist(),
              axis=alt.Axis(labelAngle=0, grid=False))

    base = alt.Chart(monthly).encode(x=x)
    bar = base.mark_bar(color="#0EA5E9", opacity=0.65, size=30).encode(
        y=alt.Y("Total_Revenue:Q", title="บาท"),
        tooltip=[alt.Tooltip("Month_Disp:O", title="เดือน"),
                 alt.Tooltip("Total_Revenue:Q", title="รายรับรวม", format=",.2f")],
    )
    line = base.mark_line(color="#0F172A", point=True, strokeWidth=2.5).encode(
        y=alt.Y("Net_Profit:Q"),
        tooltip=[alt.Tooltip("Month_Disp:O", title="เดือน"),
                 alt.Tooltip("Net_Profit:Q", title="กำไรสุทธิ", format=",.2f")],
    )
    # Data labels on top of each revenue bar.
    bar_labels = base.mark_text(
        dy=-8, color="#0284C7", fontSize=11, fontWeight="bold",
    ).encode(
        y=alt.Y("Total_Revenue:Q"),
        text=alt.Text("Total_Revenue:Q", format=",.0f"),
    )
    st.caption("🟦 แท่ง = รายรับรวม · ⬛ เส้น = กำไรสุทธิ")
    st.altair_chart(
        alt.layer(bar, line, bar_labels).resolve_scale(y="shared").properties(height=340),
        use_container_width=True,
    )


def _dashboard_monthly_table(view: pd.DataFrame) -> None:
    """Compact monthly overview (all branches): revenue / expenses / net profit."""
    monthly = (view.groupby("Month_Label", as_index=False)
                   .agg(Revenue=("Total_Revenue", "sum"),
                        Expenses=("Total_Expenses", "sum"),
                        Net=("Net_Profit", "sum"))
                   .sort_values("Month_Label"))
    monthly["Expenses"] = monthly["Expenses"].map(_negate_expense)
    monthly = monthly.rename(columns={
        "Month_Label": "เดือน", "Revenue": "รายรับรวม",
        "Expenses": "รายจ่ายรวม", "Net": "กำไรสุทธิ",
    })
    money_cols = ["รายรับรวม", "รายจ่ายรวม", "กำไรสุทธิ"]
    styler = (monthly.style
              .format(_comma, subset=money_cols)
              .apply(lambda s: [_profit_css(v) for v in s], subset=["กำไรสุทธิ"]))
    _render_styled_table(styler)


def _dashboard_table(view: pd.DataFrame) -> None:
    table = (view.groupby(["Branch_Name", "Month_Label"], as_index=False)
                 .agg(Cash=("Cash_Revenue", "sum"),
                      Transfer=("Transfer_Revenue", "sum"),
                      Water=("Water_Cost", "sum"),
                      Elec=("Electric_Cost", "sum"),
                      Rent=("Rent_Cost", "sum"),
                      Maintenance=("Maintenance_Cost", "sum"),
                      Net_Profit=("Net_Profit", "sum"))
                 .sort_values(["Branch_Name", "Month_Label"]))

    # Display expenses as negative deductions (cosmetic only — Net_Profit is
    # already revenue − expenses): Cash + Transfer − Water − Elec − Rent −
    # Maintenance = Net_Profit.
    for col in ["Water", "Elec", "Rent", "Maintenance"]:
        table[col] = table[col].map(_negate_expense)

    table = table.rename(columns={
        "Branch_Name": "สาขา", "Month_Label": "เดือน",
        "Cash": "เงินสด", "Transfer": "เงินโอน", "Water": "ค่าน้ำ",
        "Elec": "ค่าไฟ", "Rent": "ค่าเช่า", "Maintenance": "ค่าซ่อม/อื่นๆ",
        "Net_Profit": "กำไรสุทธิ",
    })
    money_cols = ["เงินสด", "เงินโอน", "ค่าน้ำ", "ค่าไฟ", "ค่าเช่า",
                  "ค่าซ่อม/อื่นๆ", "กำไรสุทธิ"]
    # pandas Styler: 'บาท'-suffixed commas via format(), green/red via apply().
    styler = (table.style
              .format(_comma, subset=money_cols)
              .apply(lambda s: [_profit_css(v) for v in s], subset=["กำไรสุทธิ"]))
    # Freeze สาขา + เดือน so they stay put while the money columns scroll.
    _render_styled_table(styler, freeze=2)


def _dashboard_branch_chart(view: pd.DataFrame) -> None:
    """Grouped-bar comparison of the selected branches, month by month."""
    metric_label = st.selectbox(
        "ตัวชี้วัดที่เปรียบเทียบ",
        ["กำไรสุทธิ", "รายรับรวม", "รายจ่ายรวม"], index=0, key="branch_metric",
    )
    metric_col = {
        "กำไรสุทธิ": "Net_Profit",
        "รายรับรวม": "Total_Revenue",
        "รายจ่ายรวม": "Total_Expenses",
    }[metric_label]

    g = (view.groupby(["Branch_Name", "Month"], as_index=False)
             .agg(Value=(metric_col, "sum"))
             .sort_values("Month"))
    if g.empty:
        return
    g["Month_Disp"] = g["Month"].dt.strftime("%b %Y")
    order = g.drop_duplicates("Month").sort_values("Month")["Month_Disp"].tolist()

    chart = alt.Chart(g).mark_bar().encode(
        x=alt.X("Month_Disp:O", title="เดือน", sort=order,
                axis=alt.Axis(labelAngle=0, grid=False)),
        xOffset=alt.XOffset("Branch_Name:N"),
        y=alt.Y("Value:Q", title="บาท"),
        color=alt.Color("Branch_Name:N", title="สาขา",
                        scale=alt.Scale(scheme="tableau10")),
        tooltip=[alt.Tooltip("Branch_Name:N", title="สาขา"),
                 alt.Tooltip("Month_Disp:O", title="เดือน"),
                 alt.Tooltip("Value:Q", title=metric_label, format=",.2f")],
    ).properties(height=360)
    st.altair_chart(chart, use_container_width=True)


def render_dashboard() -> None:
    section_header("📊", "Dashboard — ภาพรวมธุรกิจ",
                   "Revenue, expenses & net profit · filter by branch and month")
    try:
        df_master = load_and_prep_dashboard_data()
    except Exception as exc:  # noqa: BLE001
        st.error(f"Could not build the dashboard: {exc}")
        return
    if df_master.empty:
        st.info("No financial data yet. Add Sales Log / Maintenance entries first.")
        return

    branches = sorted(b for b in df_master["Branch_Name"].unique() if b)
    types = sorted(t for t in df_master["Machine_Type"].unique() if t)
    months = sorted(df_master["Month_Label"].unique())
    with st.container(border=True):
        f1, f2 = st.columns(2)
        sel_branches = f1.multiselect("Select Branch (สาขา)", branches, default=branches)
        sel_types = f2.multiselect("Select Type (ประเภทตู้ · น้ำ/น้ำแข็ง)",
                                   types, default=types)
        sel_months = st.multiselect("Select Month/Year (เดือน/ปี)", months, default=months)

    view = df_master[
        df_master["Branch_Name"].isin(sel_branches)
        & df_master["Machine_Type"].isin(sel_types)
        & df_master["Month_Label"].isin(sel_months)
    ]
    if view.empty:
        st.warning("No data matches the selected filters.")
        return

    _dashboard_kpis(view)
    st.markdown("**📈 Monthly trend · แนวโน้มรายเดือน**")
    _dashboard_trend_chart(view)
    st.markdown("**🗓️ สรุปรายเดือน (รวมทุกสาขา · หน่วย: บาท)**")
    _dashboard_monthly_table(view)
    st.markdown("**📋 สรุปแยกสาขา/เดือน (หน่วย: บาท)**")
    _dashboard_table(view)
    st.markdown("**🏢 เปรียบเทียบสาขา · รายเดือน**")
    _dashboard_branch_chart(view)


# ===========================================================================
# UI — BANK TRANSFERS (Raw_Email pending reconciliation)
# ===========================================================================
def render_bank_transfers() -> None:
    section_header("🏦", "ยอดโอนธนาคาร — ยังไม่ได้กระทบยอด",
                   "Raw_Email settlements with no matching Sales_Log period yet")
    try:
        pending = load_pending_transfers()
    except Exception as exc:  # noqa: BLE001
        st.error(f"Could not load bank transfers: {exc}")
        return
    if pending.empty:
        st.success("ทุกเดือนกระทบยอดครบแล้ว 🎉 (ไม่มียอดโอนค้างใน Raw_Email)")
        return

    months = sorted(pending["Month_Label"].unique())
    sel_months = st.multiselect("เดือน/ปี", months, default=months)
    view = pending[pending["Month_Label"].isin(sel_months)]
    if view.empty:
        st.warning("ไม่มีข้อมูลตามตัวกรองที่เลือก")
        return

    # Flat detail: one row per (Merchant, Month) showing EVERY amount column,
    # plus a grand-total row at the bottom.
    table = (view[["Merchant_Label", "Merchant_No", "Month_Label", "Last_Update",
                   *PENDING_VALUE_COLS]]
             .sort_values(["Merchant_Label", "Month_Label"]).reset_index(drop=True))
    total = {"Merchant_Label": "รวมทั้งหมด", "Merchant_No": "", "Month_Label": "",
             "Last_Update": table["Last_Update"].max()}  # overall latest transfer
    total.update({c: table[c].sum() for c in PENDING_VALUE_COLS})
    table = pd.concat([table, pd.DataFrame([total])], ignore_index=True)

    table = table.rename(columns={
        "Merchant_Label": "สาขา", "Merchant_No": "Merchant No", "Month_Label": "เดือน",
        "Last_Update": "อัพเดตล่าสุด",
        "Trans_Amount": "ยอดโอนรวม", "Commission": "ค่าธรรมเนียม", "VAT": "VAT",
        "Net_Transfer": "ยอดเข้าจริง",
    })
    money_cols = ["ยอดโอนรวม", "ค่าธรรมเนียม", "VAT", "ยอดเข้าจริง"]
    styler = table.style.format(_comma, subset=money_cols)
    st.markdown("**📋 ยอดค้างกระทบยอด · ทุกค่า (หน่วย: บาท)**")
    _render_styled_table(styler, freeze=1)
    st.caption("แสดงเฉพาะเดือน × Merchant ที่ยังไม่มีรอบบิลใน Sales_Log ครอบคลุม")


# ===========================================================================
# APP ENTRYPOINT
# ===========================================================================
def main() -> None:
    st.set_page_config(
        page_title="Koricube Console", page_icon="🧊", layout="wide"
    )
    inject_css()
    render_topbar()

    # --- Sidebar: connection status + manual data refresh -------------------
    with st.sidebar:
        st.markdown("### ⚙️ Connection")
        try:
            get_spreadsheet()  # forces auth + open; cached afterwards
            st.success("Connected · Koricube_Database")
        except Exception as exc:  # noqa: BLE001
            st.error(f"Connection failed: {exc}")
            st.stop()

        if st.button("🔄 Refresh master data", use_container_width=True):
            fetch_location_data.clear()
            fetch_raw_email_data.clear()        # bank-transfer lookups
            load_and_prep_dashboard_data.clear()  # dashboard master data
            load_pending_transfers.clear()        # pending-transfers tab
            st.rerun()

        st.markdown("---")
        st.caption("🔒 Row 1 protected · append-only writes")
        st.caption("📅 Dates · YYYY-MM-DD · 🌏 Asia/Bangkok")

    # --- Load shared master data once --------------------------------------
    try:
        location_df = fetch_location_data()
    except Exception as exc:  # noqa: BLE001
        st.error(f"Could not load Location master data: {exc}")
        st.stop()

    # --- Feature tabs · Dashboard is the FIRST (default landing) tab --------
    tab_dash, tab_bank, tab_sales, tab_maint = st.tabs(
        ["📊 Dashboard", "🏦 ยอดโอนธนาคาร", "📝 Sales Log", "🔧 Maintenance & Expenses"]
    )
    with tab_dash:
        render_dashboard()
    with tab_bank:
        render_bank_transfers()
    with tab_sales:
        render_sales_log(location_df)
    with tab_maint:
        render_maintenance(location_df)


if __name__ == "__main__":
    main()
