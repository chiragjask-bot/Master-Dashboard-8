hs
import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import re
import io
import os
import random
import zipfile
import gzip
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.comments import Comment
from openpyxl.formatting.rule import FormulaRule

# =====================================================================================
# 0. LOGIN GATE  (keeps the app private without ever putting a password on GitHub)
# =====================================================================================
# Two modes, tried in this order every time the app loads:
#   1) Streamlit secrets [auth] username/password (original method — more secure,
#      since the real credentials never touch this .py file). Used automatically
#      whenever a .streamlit/secrets.toml (or Streamlit Cloud "Secrets") [auth]
#      section is configured.
#   2) A single hardcoded ADMIN_PASSWORD gate (added per request). Used
#      automatically whenever [auth] secrets are NOT configured — e.g. for quick
#      local/admin use without setting up secrets at all.
# .streamlit/secrets.toml, if used, should look like:
#   [auth]
#   username = "your_username"
#   password = "your_password"
ADMIN_PASSWORD = "kano"

# Playful Hinglish error messages for the hardcoded ADMIN_PASSWORD path, one is
# picked at random on every wrong attempt.
ADMIN_PASSWORD_ERROR_MESSAGES = [
    "Password इल्ले! 😅 इल्ले!, खम्मा घणी भाईसा, सॉरी। तुमसे सब कुछ हो पाएगा! यहां बहुत 🤪 दिमाग मत लगाओ, इस वेबसाइट को नहीं, 😂 इस गलत पासवर्ड को छोड़ दो!",
    "❌ Password इल्ले भाईसा! 😅 इल्ले! खम्मा घणी, सॉरी। तुम बाहुबली हो, तुमसे सब कुछ हो पाएगा! पर यहाँ फालतू 🤪 दिमाग मत लगाओ। अपनी सुंदर वेबसाइट को नहीं, 😂 इस सड़े हुए गलत पासवर्ड को छोड़ दो!",
    "❌ खम्मा घणी भाईसा, Password इल्ले! 😅 sorry! तुम तो मंगल ग्रह पर पानी खोज सकते हो, तुमसे सब कुछ हो पाएगा! पर यहाँ ज़्यादा 🤪 दिमाग मत लगाओ। इस सीधे-सादे वेबसाइट को नहीं, 😂 इस जाली पासवर्ड को छोड़ दो!",
    "❌ Password इल्ले! 😅 इल्ले! खम्मा घणी भाईसा, सॉरी। लोड मत लो, तुमसे सब कुछ हो पाएगा! पर यहाँ फालतू 🤪 दिमाग मत लगाओ। दुनिया छोड़ दो, मोक्ष पकड़ लो, पर पहले 😂 इस गलत पासवर्ड को छोड़ दो!",
    "❌ अरे भाईसा! Password इल्ले! 😅 खम्मा घणी, सॉरी। तुम चाहो तो सिस्टम हिला सकते हो, तुमसे सब कुछ हो पाएगा! पर यहाँ ज़्यादा 🤪 दिमाग मत लगाओ। इस निर्दोष वेबसाइट को नहीं, 😂 इस भूतिया गलत पासवर्ड को छोड़ दो!",
]


def check_login():
    if st.session_state.get("authenticated", False):
        return True

    auth_cfg = st.secrets.get("auth", None)

    # ---- Mode 1: secrets-based username/password (unchanged from before) ----
    if auth_cfg:
        st.title("🔒 Financial File Merger & Formatter — Login")
        with st.form("login_form"):
            user = st.text_input("Username")
            pwd = st.text_input("Password", type="password")
            submitted = st.form_submit_button("Log in")
        if submitted:
            if user == auth_cfg.get("username") and pwd == auth_cfg.get("password"):
                st.session_state["authenticated"] = True
                st.rerun()
            else:
                st.error("Invalid username or password.")
        return False

    # ---- Mode 2: hardcoded ADMIN_PASSWORD gate ----
    st.markdown(
        "<p style='text-align: center; margin-top: 100px; color: Green; font-size: 18px;'>"
        "📊 Financial Data File Merger & Formatter</p>",
        unsafe_allow_html=True,
    )
    st.markdown(
        "<h1 style='text-align: center; margin-top: 0px; font-size: 20px;'>🔐 Admin Login</h1>",
        unsafe_allow_html=True,
    )

    col1, col2, col3 = st.columns([1, 1, 1])
    with col2:
        with st.form("admin_login_form"):
            pwd = st.text_input("Enter Password", type="password")
            submit = st.form_submit_button("Login", use_container_width=True)
            if submit:
                if pwd == ADMIN_PASSWORD:
                    st.session_state["authenticated"] = True
                    st.rerun()
                else:
                    st.error(random.choice(ADMIN_PASSWORD_ERROR_MESSAGES))

    dynamic_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    st.markdown(
        f"<p style='text-align: center; color: gray; font-size: 14px; margin-top: 20px;'>"
        f"Data refreshed: {dynamic_time}</p>",
        unsafe_allow_html=True,
    )
    return False


st.set_page_config(page_title="Financial File Merger & Formatter", layout="wide")

# ==========================================
# 🛡️ HIDE STREAMLIT MENU
# ==========================================
# As supplied, this sets visibility: show — i.e. the hamburger menu, header,
# toolbar, and footer all stay VISIBLE. Flip any of these to "hidden" if you
# actually want that element hidden from viewers.
hide_streamlit_ui = """
<style>
    #MainMenu {visibility: show;}
    header {visibility: show;}
    [data-testid="stToolbar"] {visibility: show;}
    footer {visibility: show;}
</style>
"""
st.markdown(hide_streamlit_ui, unsafe_allow_html=True)

# ==========================================
# 🛡️ HIDE GITHUB ICON ONLY
# ==========================================
hide_github_icon = """
<style>
    [data-testid="stToolbar"] {
        right: 2rem;
    }
    [data-testid="stToolbar"]::before {
        content: "";
    }
    button[kind="header"] {display: none;}
</style>
"""
st.markdown(hide_github_icon, unsafe_allow_html=True)

if not check_login():
    st.stop()

st.title("📊 Financial Data File Merger & Formatter")
st.markdown('<div id="main_tab"></div>', unsafe_allow_html=True)


def jump_to(anchor_id):
    """Sets the scroll target and reruns; the actual scroll happens via the
    scroll-handling snippet placed at the very end of the script."""
    st.session_state["scroll_target"] = anchor_id
    st.rerun()

# =====================================================================================
# 1. Master Sequence Profile Alignment Rule
# =====================================================================================
TAB_SEQUENCE = [
    "EQUITY_L", "SME_EQUITY_L", "Eligible_T0_Securities", "MA", "mcap", "pd", "pr", "bc", "tt",
    "BhavCopy_NSE_CM", "sec_bhavdata_full", "sec_list", "PE", "StocksTraded", "bulk",
    "mrg_trading", "CM_52_wk_High_low", "Pre-Open Market", "eq_band_changes"
]

# Regex parameters handling timestamps, dates, and alphanumeric trailing prefixes
MATCH_PATTERNS = {
    "EQUITY_L": re.compile(r"^EQUITY_L(?:_.*|\d.*)?$", re.IGNORECASE),
    "SME_EQUITY_L": re.compile(r"^SME_EQUITY_L(?:_.*|\d.*)?$", re.IGNORECASE),
    "Eligible_T0_Securities": re.compile(r"^Eligible_T0_Securities(?:_.*|\d.*)?$", re.IGNORECASE),
    "MA": re.compile(r"^MA(?:_.*|\d.*)?$", re.IGNORECASE),
    "mcap": re.compile(r"^mcap(?:_.*|\d.*)?$", re.IGNORECASE),
    "pd": re.compile(r"^pd(?:_.*|\d.*)?$", re.IGNORECASE),
    "pr": re.compile(r"^pr(?:_.*|\d.*)?$", re.IGNORECASE),
    "bc": re.compile(r"^bc(?:_.*|\d.*)?$", re.IGNORECASE),
    "tt": re.compile(r"^tt(?:_.*|\d.*)?$", re.IGNORECASE),
    "BhavCopy_NSE_CM": re.compile(r"^BhavCopy_NSE_CM(?:_.*|\d.*)?$", re.IGNORECASE),
    "sec_bhavdata_full": re.compile(r"^sec_bhavdata_full(?:_.*|\d.*)?$", re.IGNORECASE),
    "sec_list": re.compile(r"^sec_list(?:_.*|\d.*)?$", re.IGNORECASE),
    "PE": re.compile(r"^PE(?:_.*|\d.*)?$", re.IGNORECASE),
    "StocksTraded": re.compile(r"^StocksTraded(?:_.*|\d.*)?$", re.IGNORECASE),
    "bulk": re.compile(r"^bulk(?:_.*|\d.*)?$", re.IGNORECASE),
    "mrg_trading": re.compile(r"^mrg_trading(?:_.*|\d.*)?$", re.IGNORECASE),
    "CM_52_wk_High_low": re.compile(r"^CM_52_wk_High(?:_low)?(?:_.*|\d.*)?$", re.IGNORECASE),
    "Pre-Open Market": re.compile(r"^Pre-Open Market(?:_.*|\d.*)?$", re.IGNORECASE),
    "eq_band_changes": re.compile(r"^eq_band_changes(?:_.*|\d.*)?$", re.IGNORECASE),
}

# 1b. Columns that should be pre-selected (default-checked) for removal on specific tabs.
#     The person can still uncheck any of these in the UI — this only sets the starting state.
DEFAULT_REMOVE_COLUMNS = {
    "Eligible_T0_Securities": [
        "Schedule of list of securities eligible for trading in T+0 Settlement Cycle across Exchanges"
    ],
    "mcap": ["Trade Date", "Last Trade Date"],
    "pd": ["MKT", "IND_SEC", "CORP_IND"],
    "pr": ["MKT"],
    "BhavCopy_NSE_CM": ["TradDt", "BizDt", "FinInstrmTp"],
    "eq_band_changes": ["Sr. No"],
    "Pre-Open Market": ["FFM CAP", "FFM CAP (₹ Crores)"],
}


# 1b-2. Default column *order* per tab, matching the raw layout each source file
#       ships with. This only sets the sequencer's starting order — the person can
#       still drag/move columns afterward. Tabs not listed here just keep whatever
#       order the uploaded file has.
DEFAULT_COLUMN_ORDER = {
    "EQUITY_L": ["SYMBOL", "ISIN NUMBER", "NAME OF COMPANY", "SERIES", "DATE OF LISTING",
                 "MARKET LOT", "PAID UP VALUE", "FACE VALUE"],
    "SME_EQUITY_L": ["SYMBOL", "ISIN_NUMBER", "NAME_OF_COMPANY", "SERIES", "DATE_OF_LISTING",
                     "PAID_UP_VALUE", "FACE_VALUE"],
    "Eligible_T0_Securities": ["Symbol", "Name Of Company", "Series", "Effective Date"],
    "mcap": ["Symbol", "Security Name", "Series", "Face Value(Rs.)", "Market Cap(Rs.)", "Issue Size", "Close Price/Paid up value(Rs.)", "Trade Date", "Last Trade Date"],
    "pd": ["SYMBOL", "SECURITY", "SERIES", "PREV_CL_PR", "OPEN_PRICE", "HIGH_PRICE", "LOW_PRICE",
           "CLOSE_PRICE", "NET_TRDVAL", "NET_TRDQTY", "TRADES", "HI_52_WK", "LO_52_WK"],
    "pr": ["SECURITY", "PREV_CL_PR", "OPEN_PRICE", "HIGH_PRICE", "LOW_PRICE", "CLOSE_PRICE",
           "NET_TRDVAL", "NET_TRDQTY", "TRADES", "HI_52_WK", "LO_52_WK", "IND_SEC", "CORP_IND"],
    "bc": ["SYMBOL", "SECURITY", "SERIES", "PURPOSE", "RECORD_DT", "EX_DT"],
    "tt": ["SECURITY", "NET_TRDVAL", "NET_TRDQTY", "PREV_CL_PR", "CLOSE_PRIC"],
    "BhavCopy_NSE_CM": ["TckrSymb", "ISIN", "FinInstrmNm", "TtlTradgVol", "TtlTrfVal",
                        "TtlNbOfTxsExctd", "PrvsClsgPric", "ClsPric", "LastPric", "OpnPric",
                        "HghPric", "LwPric", "SttlmPric", "NewBrdLotQty", "Sgmt", "FinInstrmId",
                        "Src", "SctySrs", "SsnId"],
    "sec_bhavdata_full": ["SYMBOL", "SERIES", "DELIV_QTY", "DELIV_PER", "PREV_CLOSE",
                          "CLOSE_PRICE", "TTL_TRD_QNTY", "TURNOVER_LACS", "NO_OF_TRADES",
                          "OPEN_PRICE", "HIGH_PRICE", "LOW_PRICE", "LAST_PRICE", "AVG_PRICE", "DATE1"],
    "sec_list": ["Symbol", "Security Name", "Series", "Band", "Remarks"],
    "StocksTraded": ["Symbol", "LTP", "%chng", "Mkt Cap (₹ Crores)", "Volume (Lakhs)",
                      "Value (₹ Crores)", "Series"],
    "bulk": ["Symbol", "Security Name", "Client Name", "Buy/Sell", "Quantity Traded",
             "Trade Price / Wght. Avg. Price", "Date", "Remarks"],
    "CM_52_wk_High_low": ["SYMBOL", "Adjusted_52_Week_High", "52_Week_High_Date",
                          "Adjusted_52_Week_Low", "52_Week_Low_DT", "SERIES"],
    "eq_band_changes": ["Symbol", "Security Name", "Series", "From", "To"],
}


def default_order_for(tab, cols):
    """Reorders this tab's actual columns to match the configured default sequence,
    tolerating case/whitespace differences between the spec and the real header text.
    Any columns not covered by the default sequence keep their original relative
    order and are appended at the end, so nothing is ever silently dropped."""
    wanted = DEFAULT_COLUMN_ORDER.get(tab)
    if not wanted:
        return list(cols)
    lookup = {str(c).strip().lower(): c for c in cols}
    ordered = []
    used = set()
    for w in wanted:
        key = w.strip().lower()
        match = lookup.get(key)
        if match is not None and match not in used:
            ordered.append(match)
            used.add(match)
    for c in cols:
        if c not in used:
            ordered.append(c)
    return ordered


def default_removals_for(tab, cols):
    """Matches the configured default-removal column names against the actual
    column names on this tab, tolerating case/whitespace differences."""
    wanted = DEFAULT_REMOVE_COLUMNS.get(tab, [])
    if not wanted:
        return []
    normalized_wanted = {w.strip().lower() for w in wanted}
    return [c for c in cols if str(c).strip().lower() in normalized_wanted]


# 1c. Default "start cell" for each tab — this is the single source of truth for
#     both (a) the "Start cell for {tab}" data-crop input's default value, and
#     (b) which cell the tab opens/scrolls to in Excel. Anything not listed here
#     defaults to blank / "A1" ("zero zero" — the normal top-left start).
DEFAULT_START_CELLS = {
    "MA": "B201",
    "Eligible_T0_Securities": "B3",
    "mrg_trading": "A11",
}


def get_default_view_cell(tab):
    return DEFAULT_START_CELLS.get(tab, "A1")


# 1d. Row-freeze note: every tab's header actually lands at the same row (row 2 —
#     right after the "⬆️ Main Tab" nav row), regardless of the tab's own start
#     cell, so freezing right after the header row (done generically below) is
#     already correct for every tab. No per-tab override needed.
CUSTOM_FREEZE_ROWS = {}

# 1e. Which cell the "⬆️ Main Tab" jump-back link is written to, per tab.
#     Defaults to A1; a couple of tabs were reported as needing it at B1 instead.
NAV_LINK_CELL_OVERRIDES = {
    "Eligible_T0_Securities": "B1",
    "MA": "B1",
    "mrg_trading": "B1",
}


def get_nav_link_cell(tab):
    return NAV_LINK_CELL_OVERRIDES.get(tab, "A1")


def render_column_sequencer(state_key, current_columns, allow_delete=False, protected=None,
                             label="Column order", default_order=None):
    """Renders a pick + ◀ Move Left / Move Right ▶ (+ optional 🗑 Delete) control.
    Order (and deletions) persist in st.session_state[state_key] across reruns.
    default_order: optional starting sequence (e.g. from DEFAULT_COLUMN_ORDER) used
    only the first time this sequencer initializes; the person can still reorder
    freely afterward. Falls back to current_columns as-is when not provided.
    Returns the ordered list of column names to use for output.
    """
    protected = protected or []

    if state_key not in st.session_state:
        st.session_state[state_key] = list(default_order) if default_order else list(current_columns)
    order = st.session_state[state_key]

    # Keep in sync with the current column set: drop stale entries, append new ones.
    order = [c for c in order if c in current_columns]
    for c in current_columns:
        if c not in order:
            order.append(c)
    st.session_state[state_key] = order

    if not order:
        return order

    st.caption(f"🔀 {label} — pick a column, then move it left/right" + (" or delete it:" if allow_delete else ":"))
    pick_col, left_col, right_col, del_col = st.columns([3, 1, 1, 1])
    with pick_col:
        pick = st.selectbox(label, options=order, key=f"{state_key}_pick", label_visibility="collapsed")
    with left_col:
        if st.button("◀ Left", key=f"{state_key}_left"):
            i = order.index(pick)
            if i > 0:
                order[i - 1], order[i] = order[i], order[i - 1]
                st.session_state[state_key] = order
                st.rerun()
    with right_col:
        if st.button("Right ▶", key=f"{state_key}_right"):
            i = order.index(pick)
            if i < len(order) - 1:
                order[i + 1], order[i] = order[i], order[i + 1]
                st.session_state[state_key] = order
                st.rerun()
    if allow_delete:
        with del_col:
            disabled = pick in protected
            if st.button("🗑 Delete", key=f"{state_key}_del", disabled=disabled,
                         help="This column is required and can't be deleted" if disabled else None):
                order.remove(pick)
                st.session_state[state_key] = order
                st.rerun()

    st.caption(" → ".join(order))

    # ---- Additional sequence box: type the exact order, comma-separated ----
    type_col, apply_col = st.columns([5, 1])
    with type_col:
        typed = st.text_input(
            "Or type the exact sequence here (comma-separated column names), then Apply:",
            key=f"{state_key}_typed",
            placeholder=", ".join(order),
        )
    with apply_col:
        st.write("")
        apply_clicked = st.button("Apply", key=f"{state_key}_apply")

    if apply_clicked:
        typed_names = [t.strip() for t in typed.split(",") if t.strip()]
        if not typed_names:
            st.warning("Type at least one column name before clicking Apply.")
        else:
            lookup = {c.strip().lower(): c for c in order}
            matched, unmatched = [], []
            for name in typed_names:
                actual = lookup.get(name.strip().lower())
                if actual and actual not in matched:
                    matched.append(actual)
                elif not actual:
                    unmatched.append(name)

            if not allow_delete:
                # Non-deletable sequencers: anything left unmentioned is appended
                # at the end rather than dropped, so no data silently disappears.
                for c in order:
                    if c not in matched:
                        matched.append(c)
            else:
                # Deletable sequencers: protected columns are kept even if the
                # person forgot to type them; everything else not typed is dropped.
                for c in protected:
                    if c in order and c not in matched:
                        matched.append(c)

            if unmatched:
                st.warning(f"Not found (ignored): {', '.join(unmatched)}")
            if matched:
                st.session_state[state_key] = matched
                st.rerun()

    return order


# 2. Strict Custom Numeric Formatting Strings Configuration
NUMBER_FORMATS = {
    'price': '#,##0.00',
    'qty': '#,##0',
    'date': 'DD-MMM-YYYY',  # 4-digit year everywhere
    'percent': '0.00"%"',
    'ratio': '0.00',
    'crores': '#,##0.00" Cr"',
    'lakhs': '#,##0.00" L"',
    'number': '0',
    'text': '@'
}


# =====================================================================================
# 2b. Master_Dashboard-8 builder — joins every consolidated tab on Symbol into one
#     wide reference sheet. Runs automatically at the end of "Execute Structural
#     Consolidation", right after all tabs are written into the same workbook.
# =====================================================================================
MASTER_SHEET_NAME = "Master_Dashboard-8"
MASTER_HIGHLIGHT_COLOR = "EAD1DC"
MASTER_HEADER_SCAN_ROWS = 15
MASTER_SYMBOL_ALIASES = ["SYMBOL", "TckrSymb", "Symb", "Symbol"]
# Safety cap: one Data Validation object is created per row for the Symbol
# "box display" quick-view tooltip. Benchmarked at ~6,000 rows in well under a
# second and a tiny file-size increase, so this ceiling is just a sane backstop,
# not a realistic limit for NSE-sized symbol lists.
MAX_BOX_DISPLAY_ROWS = 20000

MASTER_FIELD_MAP = [
    {"label": "Symbol", "sheet": "BhavCopy_NSE_CM", "aliases": ["TckrSymb", "SYMBOL", "Symb"], "format": "text", "isKey": True},
    # Pinned right next to Symbol (not at the far end with the other hyperlink
    # columns) — feature request: a clickable link "dot" visible without
    # scrolling, right beside the symbol name. Freeze panes below is widened
    # to column C to keep both Symbol and this dot on screen while scrolling.
    {"label": "NSE Chart", "sheet": None, "aliases": [], "format": "text"},
    {"label": "ISIN", "sheet": "BhavCopy_NSE_CM", "aliases": ["ISIN", "ISIN NUMBER"], "format": "text"},
    {"label": "Series", "sheet": "BhavCopy_NSE_CM", "aliases": ["SctySrs", "SERIES", "Series", "Srs"], "format": "text"},
    {"label": "Company Name (Capital)", "sheet": "BhavCopy_NSE_C