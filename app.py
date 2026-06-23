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
from datetime import datetime, date
from html import escape
from typing import Any, List, Optional

import pandas as pd
import streamlit as st

# pdfplumber is only needed for the bank-statement feature; import lazily-safe
# so the rest of the app still loads if the wheel is missing on a given host.
try:
    import pdfplumber
    _PDFPLUMBER_AVAILABLE = True
except Exception:  # pragma: no cover - environment dependent
    _PDFPLUMBER_AVAILABLE = False

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

# Maintenance defaults.
ERROR_CODES = ["E7", "A05", "Other"]
STATUS_PENDING = "รอซ่อม"  # default ticket status

# Buddhist Era offset (BE = AD + 543).
BE_OFFSET = 543

# ``USER_ENTERED`` lets Google Sheets store our numeric strings as real numbers
# and our unambiguous ISO ``YYYY-MM-DD`` strings as real dates, so the row-1
# ArrayFormula logic can compute on them directly.
VALUE_INPUT_OPTION = "USER_ENTERED"


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

    Uses a service-account key file as mandated by the architecture.
    """
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
# PDF EXTRACTION (Bank settlement report)
# ===========================================================================
# Columns we expect to recover from each settlement row, in payload order.
RAW_EMAIL_FIELDS = [
    "Transfer_Date",
    "Merchant_No",
    "Trans_Amount",
    "Commission",
    "VAT",
    "Net_Transfer",
]


def extract_settlement_rows(pdf_file) -> pd.DataFrame:
    """
    Extract settlement rows from a bank PDF using pdfplumber.

    Strategy
    --------
    1. Pull every table from every page with ``extract_tables()``.
    2. Heuristically locate the 6 columns of interest.
    3. Normalise dates (incl. Buddhist-Era -> AD) and clean numerics.

    The whole body is wrapped by the caller in try/except so a structural
    change in the bank's layout surfaces as a friendly error, not a crash.

    Returns a DataFrame with exactly ``RAW_EMAIL_FIELDS`` columns.
    """
    if not _PDFPLUMBER_AVAILABLE:
        raise RuntimeError(
            "pdfplumber is not installed in this environment "
            "(`pip install pdfplumber`)."
        )

    collected: List[List[Any]] = []

    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            for table in page.extract_tables() or []:
                for row in table:
                    cells = [("" if c is None else str(c).strip()) for c in row]
                    if not any(cells):
                        continue

                    # Skip obvious header rows.
                    joined = " ".join(cells).lower()
                    if any(k in joined for k in ("merchant", "date", "amount",
                                                 "commission", "vat", "net",
                                                 "วันที่", "ยอด")):
                        # Header-like; do not treat as data.
                        continue

                    parsed = _parse_settlement_cells(cells)
                    if parsed is not None:
                        collected.append(parsed)

    df = pd.DataFrame(collected, columns=RAW_EMAIL_FIELDS)
    if df.empty:
        raise ValueError(
            "No settlement rows could be located in this PDF. "
            "The layout may have changed — please verify the file."
        )
    return df


def _parse_settlement_cells(cells: List[str]) -> Optional[List[Any]]:
    """
    Map a raw list of table cells onto the 6 target fields.

    Expects, after dropping empties, a row resembling:
        [Transfer_Date, Merchant_No, Trans_Amount, Commission, VAT, Net_Transfer]

    Returns ``None`` for rows that clearly are not data (so callers can skip).
    """
    values = [c for c in cells if c != ""]
    if len(values) < 6:
        return None

    # Take the first cell as the date and the last four as the money columns;
    # the merchant number is whatever sits between the date and the amounts.
    transfer_date = normalize_date_string(values[0])

    # The four right-most numeric-looking cells = amount/commission/vat/net.
    trans_amount = clean_number(values[-4])
    commission = clean_number(values[-3])
    vat = clean_number(values[-2])
    net_transfer = clean_number(values[-1])

    # Merchant number: first non-date token that is not one of the 4 money cells.
    merchant_no = ""
    for v in values[1:-4]:
        digits = re.sub(r"\D", "", v)
        if digits:
            merchant_no = digits
            break
    if not merchant_no:
        merchant_no = re.sub(r"\D", "", values[1]) if len(values) > 1 else ""

    # If the money columns failed to parse, this almost certainly isn't a data
    # row — let the caller skip it.
    if None in (trans_amount, commission, vat, net_transfer):
        return None

    return [transfer_date, merchant_no, trans_amount, commission, vat, net_transfer]


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
        .block-container{ padding-top:1.4rem; padding-bottom:3rem; max-width:1160px; }
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

        /* ---- Mobile ---- */
        @media (max-width:640px){
            .block-container{ padding-left:.7rem; padding-right:.7rem; }
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


def render_reconciliation_cue(web_total: float, cash_collected: float) -> None:
    """Live colour-coded delta between cash counted and web telemetry total."""
    diff = round(cash_collected - web_total, 2)
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

    labels = location_df.apply(machine_label, axis=1).tolist()
    label_to_index = {label: idx for idx, label in enumerate(labels)}

    # Selectbox stays OUTSIDE a form: selecting a machine reruns instantly so the
    # period fields and the live reconciliation cue can react in real time.
    selected_label = st.selectbox("Select Machine (เลือกตู้)", labels)
    machine = location_df.iloc[label_to_index[selected_label]]

    payment_type = machine[LOC_PAYMENT_TYPE]
    is_cash_only = payment_type == PAYMENT_CASH_ONLY

    # Compact, single-line status badge (replaces bulky headings).
    render_status_card(machine, payment_type, is_cash_only)

    # All inputs grouped inside one clean white card.
    with st.container(border=True):
        # CRITICAL: period inputs only appear for 'โอน+เงินสด' machines.
        period_start_obj: Optional[date] = None
        period_end_obj: Optional[date] = None
        if not is_cash_only:
            st.markdown("**🗓️ Reconciliation period · ช่วงรอบบิล**")
            p1, p2 = st.columns(2)
            period_start_obj = p1.date_input("Period Start", value=None, key="sl_pstart")
            period_end_obj = p2.date_input("Period End", value=None, key="sl_pend")

        # CRITICAL: Web Total is available for EVERY machine (telemetry audit).
        st.markdown("**💰 Amounts · ยอดเงิน**")
        m1, m2 = st.columns(2)
        web_total = m1.number_input(
            "Web Total · ยอดเว็บ (฿)", min_value=0.0, step=1.0, format="%.2f", key="sl_web"
        )
        cash_collected = m2.number_input(
            "Cash Collected · เงินสดที่เก็บได้ (฿)",
            min_value=0.0, step=1.0, format="%.2f", key="sl_cash",
        )

        # Live, colour-coded reconciliation feedback.
        render_reconciliation_cue(web_total, cash_collected)

        remark = st.text_area(
            "Remark · หมายเหตุ", placeholder="Optional notes…", key="sl_remark"
        )

        submitted = st.button("Submit Sales Log", type="primary",
                              use_container_width=True, key="sl_submit")

    if not submitted:
        return

    # Re-stamp the CURRENT master values at submit time (per spec).
    branch_name = machine[LOC_BRANCH]
    merchant_no = machine[LOC_MERCHANT_NO]
    machines_shared = machine[LOC_MACHINES_SHARED]

    period_start = "" if is_cash_only else date_to_iso(period_start_obj)
    period_end = "" if is_cash_only else date_to_iso(period_end_obj)

    # Validate period inputs for transfer machines.
    if not is_cash_only and (not period_start or not period_end):
        st.error("Please provide both Period Start and Period End for this machine.")
        return
    if period_start and period_end and period_start > period_end:
        st.error("Period Start must not be after Period End.")
        return

    # Payload order is contractually fixed — do not reorder.
    payload = [
        today_iso(),              # Collection_Date
        machine[LOC_MACHINE_ID],  # Machine_ID
        branch_name,              # Branch_Name
        merchant_no,              # Merchant_No
        payment_type,             # Payment_Type
        machines_shared,          # Machines_Shared
        float(web_total),         # Web_Total
        float(cash_collected),    # Cash_Collected
        period_start,             # Period_Start
        period_end,               # Period_End
        remark.strip(),           # Remark
    ]

    try:
        append_row_safe(WS_SALES_LOG, payload)
    except Exception as exc:  # noqa: BLE001 - present any backend error nicely
        st.error(f"Failed to write to Sales_Log: {exc}")
        return

    st.success(f"✅ Sales log saved for {machine[LOC_MACHINE_ID]} ({branch_name}).")
    st.caption(f"Collection date stamped: {payload[0]}")


# ===========================================================================
# UI — FEATURE 2: BANK STATEMENT PROCESSOR  (อ่าน PDF ด้วย pdfplumber)
# ===========================================================================
def render_pdf_processor() -> None:
    section_header("🏦", "Bank Statement Processor — อ่าน PDF ด้วย pdfplumber",
                   "Extract settlement rows · auto-convert Buddhist Era → AD")

    if not _PDFPLUMBER_AVAILABLE:
        st.error("pdfplumber is not available in this environment.")
        return

    uploaded = st.file_uploader(
        "Upload bank settlement report (PDF)", type=["pdf"],
        accept_multiple_files=False,
    )
    if uploaded is None:
        st.info("Upload a settlement PDF to begin extraction.")
        return

    # Robust extraction — a layout change yields a graceful error, never a crash.
    try:
        df = extract_settlement_rows(uploaded)
    except Exception as exc:  # noqa: BLE001
        st.error(
            "Could not parse this PDF. The bank's layout may have changed.\n\n"
            f"Details: {exc}"
        )
        return

    st.success(f"Extracted {len(df)} settlement row(s). Review before saving.")
    # Let the operator review (and lightly fix) before committing to the sheet.
    edited = st.data_editor(
        df, num_rows="dynamic", use_container_width=True, key="pdf_editor"
    )

    if st.button("Append all rows to Raw_Email", type="primary",
                 use_container_width=True):
        success, failed = 0, 0
        for _, row in edited.iterrows():
            payload = [
                normalize_date_string(row["Transfer_Date"]),  # Transfer_Date (BE->AD)
                re.sub(r"\D", "", str(row["Merchant_No"])),   # Merchant_No
                clean_number(row["Trans_Amount"]),            # Trans_Amount
                clean_number(row["Commission"]),              # Commission
                clean_number(row["VAT"]),                     # VAT
                clean_number(row["Net_Transfer"]),            # Net_Transfer
            ]
            try:
                append_row_safe(WS_RAW_EMAIL, payload)
                success += 1
            except Exception:  # noqa: BLE001
                failed += 1

        if success:
            st.success(f"✅ Appended {success} row(s) to Raw_Email.")
        if failed:
            st.error(f"⚠️ {failed} row(s) failed to append.")


# ===========================================================================
# UI — FEATURE 3: MAINTENANCE FORM  (แจ้งซ่อม)
# ===========================================================================
def render_maintenance(location_df: pd.DataFrame) -> None:
    section_header("🔧", "Maintenance — แจ้งซ่อม",
                   "Log a repair ticket · default status รอซ่อม")

    if location_df.empty:
        st.warning("No machines found in the Location master sheet.")
        return

    labels = location_df.apply(machine_label, axis=1).tolist()
    label_to_index = {label: idx for idx, label in enumerate(labels)}

    with st.form("maintenance_form", clear_on_submit=True):
        top1, top2 = st.columns([2, 1])
        selected_label = top1.selectbox("Select Machine (เลือกตู้)", labels)
        error_code = top2.selectbox("Error Code", ERROR_CODES)
        issue_desc = st.text_area("Issue Description (รายละเอียดปัญหา)")
        repair_cost = st.number_input(
            "Repair Cost · ค่าซ่อม (฿)", min_value=0.0, step=1.0, format="%.2f"
        )
        submitted = st.form_submit_button(
            "Submit Repair Ticket", type="primary", use_container_width=True
        )

    if not submitted:
        return

    machine = location_df.iloc[label_to_index[selected_label]]

    # Payload order is contractually fixed.
    payload = [
        today_iso(),               # Report_Date
        machine[LOC_MACHINE_ID],   # Machine_ID
        error_code,                # Error_Code
        issue_desc.strip(),        # Issue_Desc
        float(repair_cost),        # Repair_Cost
        "",                        # Resolved_Date (blank until resolved)
        STATUS_PENDING,            # Status -> "รอซ่อม"
    ]

    try:
        append_row_safe(WS_MAINTENANCE, payload)
    except Exception as exc:  # noqa: BLE001
        st.error(f"Failed to write to Maintenance: {exc}")
        return

    st.success(
        f"✅ Repair ticket created for {machine[LOC_MACHINE_ID]} "
        f"(status: {STATUS_PENDING})."
    )


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

    # --- Feature tabs -------------------------------------------------------
    tab_sales, tab_pdf, tab_maint = st.tabs(
        ["📋 Sales Log", "🏦 Bank Statement", "🔧 Maintenance"]
    )
    with tab_sales:
        render_sales_log(location_df)
    with tab_pdf:
        render_pdf_processor()
    with tab_maint:
        render_maintenance(location_df)


if __name__ == "__main__":
    main()
