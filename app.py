import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import re
import io
import os
import zipfile
import gzip
from datetime import datetime
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# =====================================================================================
# 0. OPTIONAL LOGIN GATE  (keeps the app private without ever putting a password on GitHub)
# =====================================================================================
# How this works:
#   - Credentials live ONLY in Streamlit secrets (a file called .streamlit/secrets.toml)
#   - That file is added to .gitignore, so it is never pushed to GitHub.
#   - On Streamlit Community Cloud, you paste the same values into
#     "App settings -> Secrets" in the web UI instead of a file.
#
# .streamlit/secrets.toml should look like:
#   [auth]
#   username = "your_username"
#   password = "your_password"
#
# If no [auth] secrets are configured, the login gate is skipped automatically
# (useful for local testing) so the app still runs.
def check_login():
    auth_cfg = st.secrets.get("auth", None)
    if not auth_cfg:
        return True  # no credentials configured -> don't block access

    if st.session_state.get("authenticated", False):
        return True

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


st.set_page_config(page_title="Financial File Merger & Formatter", layout="wide")

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
    "pd": ["MKT", "IND_SEC", "CORP_IND"],
    "pr": ["MKT"],
    "BhavCopy_NSE_CM": ["TradDt", "BizDt", "FinInstrmTp"],
    "eq_band_changes": ["Sr. No"],
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
    "pd": ["SYMBOL", "SECURITY", "SERIES", "PREV_CL_PR", "OPEN_PRICE", "HIGH_PRICE", "LOW_PRICE",
           "CLOSE_PRICE", "NET_TRDVAL", "NET_TRDQTY", "TRADES", "HI_52_WK", "LO_52_WK"],
    "pr": ["SECURITY", "PREV_CL_PR", "OPEN_PRICE", "HIGH_PRICE", "LOW_PRICE", "CLOSE_PRICE",
           "NET_TRDVAL", "NET_TRDQTY", "TRADES", "HI_52_WK", "LO_52_WK", "IND_SEC", "CORP_IND"],
    "bc": ["SYMBOL", "SECURITY", "SERIES", "PURPOSE", "RECORD_DT", "EX_DT"],
    "tt": ["SECURITY", "NET_TRDVAL", "NET_TRDQTY", "PREV_CL_PR", "CLOSE_PRIC"],
    "BhavCopy_NSE_CM": ["ISIN", "TckrSymb", "FinInstrmNm", "TtlTradgVol", "TtlTrfVal",
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
    "MA": "B9",
    "Eligible_T0_Securities": "B3",
    "mrg_trading": "A11",
}


def get_default_view_cell(tab):
    return DEFAULT_START_CELLS.get(tab, "A1")


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

MASTER_FIELD_MAP = [
    {"label": "Symbol", "sheet": "BhavCopy_NSE_CM", "aliases": ["TckrSymb", "SYMBOL", "Symb"], "format": "text", "isKey": True},
    {"label": "ISIN", "sheet": "BhavCopy_NSE_CM", "aliases": ["ISIN", "ISIN NUMBER"], "format": "text"},
    {"label": "Series", "sheet": "BhavCopy_NSE_CM", "aliases": ["SctySrs", "SERIES", "Series", "Srs"], "format": "text"},
    {"label": "Company Name (Capital)", "sheet": "BhavCopy_NSE_CM",
     "aliases": ["FinInstrmNm", "NAME OF COMPANY", "Name Of Company", "Security Name", "SECURITY", "Security",
                 "COMPANY NAME", "COMPANY'S NAME", "Company Name", "Company's Name"], "format": "text"},
    {"label": "Company Name", "sheet": "EQUITY_L",
     "aliases": ["NAME OF COMPANY", "Name Of Company", "Security Name", "SECURITY", "Security",
                 "COMPANY NAME", "COMPANY'S NAME", "Company Name", "Company's Name"], "format": "text"},
    {"label": "Date of Listing", "sheet": "EQUITY_L", "aliases": ["DATE OF LISTING"], "format": "date"},
    {"label": "Trade Date", "sheet": "BhavCopy_NSE_CM", "aliases": ["TradDt", "Trade Date"], "format": "date"},
    {"label": "Segment", "sheet": "BhavCopy_NSE_CM", "aliases": ["Src"], "format": "text"},
    {"label": "Delivery %", "sheet": "sec_bhavdata_full",
     "aliases": ["DELIV PER", "DELIV %", "delivery percentage", "Delivery Percentage (%)", "DELIV_PER"], "format": "percent"},
    {"label": "% Change", "sheet": "StocksTraded", "aliases": ["%chng", "% Change"], "format": "percent"},
    {"label": "Close Price", "sheet": "BhavCopy_NSE_CM", "aliases": ["ClsPric", "CLOSE PRICE", "Close Price", "CLOSE_PRICE"], "format": "price"},
    {"label": "CMP/LTP", "sheet": "BhavCopy_NSE_CM", "aliases": ["LastPric", "LAST PRICE", "Last Price", "LTP", "LAST_PRICE"], "format": "price"},
    {"label": "Prev Close", "sheet": "BhavCopy_NSE_CM", "aliases": ["PrvsClsgPric", "PREV CLOSE", "Previous close", "PREV_CL_PR", "PREV_CLOSE"], "format": "price"},
    {"label": "Open (Rs.)", "sheet": "BhavCopy_NSE_CM", "aliases": ["OpnPric", "Open Price", "OPEN PRICE", "OPEN_PRICE"], "format": "price"},
    {"label": "High (Rs.)", "sheet": "BhavCopy_NSE_CM", "aliases": ["HghPric", "HIGH PRICE", "High Price", "HIGH_PRICE"], "format": "price"},
    {"label": "Low (Rs.)", "sheet": "BhavCopy_NSE_CM", "aliases": ["LwPric", "Low Price", "LOW PRICE", "LOW_PRICE"], "format": "price"},
    {"label": "Turnover (Rs.)", "sheet": "BhavCopy_NSE_CM",
     "aliases": ["TtlTrfVal", "NET_TRDVAL", "NET_TRD_VAL", "NET TRD VAL", "NET TRDVAL", "Turnover (Rs.)", "NET TRADED VALUE", "Net Traded Value", "Traded Value"], "format": "qty"},
    {"label": "Traded Qty", "sheet": "BhavCopy_NSE_CM",
     "aliases": ["TtlTradgVol", "TTL TRD QNTY", "TRADED QUANTITY", "NET_TRDQTY", "Traded Qty", "NET TRD QTY", "NET TRDQTY", "TTL_TRD_QNTY"], "format": "qty"},
    {"label": "No. of Trades", "sheet": "BhavCopy_NSE_CM", "aliases": ["TtlNbOfTxsExctd", "No. of Trades", "NO OF TRADES", "TRADES", "Trade", "NO_OF_TRADES"], "format": "qty"},
    {"label": "Market Lot", "sheet": "BhavCopy_NSE_CM", "aliases": ["NewBrdLotQty", "MARKET LOT", "Market Lot"], "format": "qty"},
    {"label": "Volume (Lakhs)", "sheet": "StocksTraded", "aliases": ["Volume (Lakhs)"], "format": "lakhs"},
    {"label": "Value (Rs. Crores)", "sheet": "StocksTraded", "aliases": ["Value (Rs Crores)", "Value (\u20b9 Crores)"], "format": "crores"},
    {"label": "Mkt Cap (Rs. Crores)", "sheet": "StocksTraded", "aliases": ["Mkt Cap (Rs Crores)", "Mkt Cap (\u20b9 Crores)", "Market Cap (\u20b9 Crores)"], "format": "crores"},
    {"label": "Market Cap (Rs.)", "sheet": "mcap", "aliases": ["Market Cap(Rs.)"], "format": "qty"},
    {"label": "Issue Size", "sheet": "mcap", "aliases": ["Issue Size"], "format": "qty"},
    {"label": "Category", "sheet": "mcap", "aliases": ["Category"], "format": "text"},
    {"label": "Face Value", "sheet": "EQUITY_L", "aliases": ["FACE VALUE", "Face Value(Rs.)"], "format": "price"},
    {"label": "Delivery Qty", "sheet": "sec_bhavdata_full",
     "aliases": ["DELIV QTY", "DELIV QUANTITY", "Delivery quantity", "DELIVERY QNTY", "DELIV_QNTY", "DELIV QNTY", "DELIV_QTY"], "format": "qty"},
    {"label": "52W High", "sheet": "CM_52_wk_High_low",
     "aliases": ["Adjusted_52_Week_High", "52_Week_High", "52W_High", "52 Week High", "52W High", "HI_52_WK"], "format": "price"},
    {"label": "52W High Date", "sheet": "CM_52_wk_High_low",
     "aliases": ["52_Week_High_Date", "52 Week High Date", "52_Week_High_DT", "52W High Date", "52 W High Date", "52 W High Dt.", "52W High Dt."], "format": "date"},
    {"label": "52W Low", "sheet": "CM_52_wk_High_low",
     "aliases": ["Adjusted_52_Week_Low", "52 Week Low", "52_Week_Low", "52W_Low", "52W Low", "LO_52_WK"], "format": "price"},
    {"label": "52W Low Date", "sheet": "CM_52_wk_High_low",
     "aliases": ["52_Week_Low_DT", "52 Week Low Date", "52 W Low Date", "52 W Low Dt.", "52W Low Dt."], "format": "date"},
    {"label": "Symbol P/E", "sheet": "PE", "aliases": ["SYMBOL P/E", "Symbol P/E"], "format": "ratio"},
    {"label": "Adjusted P/E", "sheet": "PE", "aliases": ["ADJUSTED P/E", "Adjusted P/E"], "format": "ratio"},
    {"label": "T0 Tag", "sheet": "Eligible_T0_Securities", "aliases": ["SERIES", "SctySrs", "Srs", "Series"], "format": "text"},
    {"label": "T0 Effective Date", "sheet": "Eligible_T0_Securities", "aliases": ["Effective Date"], "format": "text"},
    {"label": "Band", "sheet": "sec_list", "aliases": ["Band"], "format": "number"},
    {"label": "Remarks", "sheet": "sec_list", "aliases": ["Remarks"], "format": "text"},
    {"label": "Paid Up Value", "sheet": "EQUITY_L", "aliases": ["PAID UP VALUE"], "format": "price"},
]

MASTER_NUMBER_FORMATS = {
    "price": "#,##0.00",
    "qty": "#,##0",
    "date": "dd-mmm-yyyy",
    "percent": '0.00"%"',
    "ratio": "0.00",
    "crores": '#,##0.00" Cr"',
    "lakhs": '#,##0.00" L"',
    "number": "0",
    "text": "@",
}


def md_normalize_header(text) -> str:
    if text is None:
        return ""
    return " ".join(str(text).strip().lower().split())


def md_find_header_row(ws, all_aliases_norm: set, scan_rows: int = MASTER_HEADER_SCAN_ROWS) -> int:
    max_row = min(scan_rows, ws.max_row)
    max_col = ws.max_column
    if max_row == 0 or max_col == 0:
        return -1
    best_row, best_score = -1, 0
    for r in range(1, max_row + 1):
        score = 0
        for c in range(1, max_col + 1):
            norm = md_normalize_header(ws.cell(row=r, column=c).value)
            if norm and norm in all_aliases_norm:
                score += 1
        if score > best_score:
            best_score, best_row = score, r
    return best_row if best_score > 0 else -1


def md_build_header_index(ws, header_row: int) -> dict:
    idx = {}
    for c in range(1, ws.max_column + 1):
        norm = md_normalize_header(ws.cell(row=header_row, column=c).value)
        if norm:
            idx[norm] = c
    return idx


def md_match_column(header_index: dict, aliases: list) -> int:
    for a in aliases:
        norm = md_normalize_header(a)
        if norm in header_index:
            return header_index[norm]
    return -1


def md_build_master_dashboard(wb):
    """wb is an openpyxl Workbook already holding the freshly consolidated tabs
    (values, not formulas — safe to read cell.value directly, no data_only reload needed)."""
    all_aliases = list(MASTER_SYMBOL_ALIASES)
    for f in MASTER_FIELD_MAP:
        all_aliases += f["aliases"]
    all_aliases_norm = {md_normalize_header(a) for a in all_aliases}

    fields_by_sheet = {}
    for f in MASTER_FIELD_MAP:
        fields_by_sheet.setdefault(f["sheet"], []).append(f)

    master_data = {}
    symbol_order = []
    log = []

    for sheet_name, fields in fields_by_sheet.items():
        if sheet_name not in wb.sheetnames:
            log.append(f'Master_Dashboard-8: "{sheet_name}" tab not present in this workbook — skipped.')
            continue
        ws = wb[sheet_name]
        header_row = md_find_header_row(ws, all_aliases_norm)
        if header_row == -1:
            log.append(f'Master_Dashboard-8: could not detect a header row on "{sheet_name}" — skipped.')
            continue
        header_index = md_build_header_index(ws, header_row)
        symbol_col = md_match_column(header_index, MASTER_SYMBOL_ALIASES)
        if symbol_col == -1:
            log.append(f'Master_Dashboard-8: no Symbol-like column found on "{sheet_name}" — skipped.')
            continue
        field_cols = [md_match_column(header_index, f["aliases"]) for f in fields]

        for r in range(header_row + 1, ws.max_row + 1):
            symbol_raw = ws.cell(row=r, column=symbol_col).value
            if symbol_raw is None or str(symbol_raw).strip() == "":
                continue
            symbol = str(symbol_raw).strip()
            if symbol not in master_data:
                master_data[symbol] = {}
                symbol_order.append(symbol)

            for f, col in zip(fields, field_cols):
                label = f["label"]
                if f.get("isKey"):
                    master_data[symbol][label] = symbol
                    continue
                if col == -1:
                    continue
                val = ws.cell(row=r, column=col).value
                cur = master_data[symbol].get(label)
                if cur is None or cur == "":
                    master_data[symbol][label] = val

    symbol_order = sorted(symbol_order)
    labels = [f["label"] for f in MASTER_FIELD_MAP]
    rows = [[master_data[s].get(l, "") for l in labels] for s in symbol_order]
    df = pd.DataFrame(rows, columns=labels)
    return df, log


def md_write_master_sheet(wb, df, column_order=None):
    """Adds/overwrites Master_Dashboard-8 directly on the same workbook object.
    column_order: optional list of field labels (subset/reordered) controlling
    which columns appear and in what order. Defaults to the full MASTER_FIELD_MAP."""
    if MASTER_SHEET_NAME in wb.sheetnames:
        del wb[MASTER_SHEET_NAME]
    ws = wb.create_sheet(MASTER_SHEET_NAME)

    field_lookup = {f["label"]: f for f in MASTER_FIELD_MAP}
    labels = [l for l in (column_order or list(df.columns)) if l in field_lookup]
    if not labels:
        labels = list(df.columns)
    df = df[labels]
    formats = [MASTER_NUMBER_FORMATS.get(field_lookup[l]["format"], "@") for l in labels]

    header_fill = PatternFill(start_color=MASTER_HIGHLIGHT_COLOR, end_color=MASTER_HIGHLIGHT_COLOR, fill_type="solid")
    bold_font = Font(name="Arial", bold=True)
    body_font = Font(name="Arial")
    thin = Side(style="thin", color="CCCCCC")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for c, label in enumerate(labels, start=1):
        cell = ws.cell(row=1, column=c, value=label)
        cell.fill = header_fill
        cell.font = bold_font
        cell.alignment = Alignment(horizontal="center")
        cell.border = border

    for r, row in enumerate(df.itertuples(index=False), start=2):
        for c, val in enumerate(row, start=1):
            v = None if (val == "" or pd.isna(val)) else val
            cell = ws.cell(row=r, column=c, value=v)
            cell.font = body_font
            cell.number_format = formats[c - 1]
            cell.border = border

    for c in range(1, len(labels) + 1):
        ws.column_dimensions[ws.cell(row=1, column=c).column_letter].width = 20
    ws.freeze_panes = "A2"
    # Move it to the front so it's the first thing a reviewer sees, right after Main Tab.
    wb.move_sheet(MASTER_SHEET_NAME, offset=-(len(wb.sheetnames) - 2))
    return wb


def resolve_tab_name(filename):
    """Bypasses custom timestamps to cleanly match dynamic files to structural tab sequences."""
    base_name = filename.rsplit('.', 1)[0]
    # strip known compression suffixes that can appear before the real extension
    for comp_ext in ('.csv', '.xls', '.xlsx'):
        if base_name.lower().endswith(comp_ext):
            base_name = base_name[: -len(comp_ext)]
    for tab_name, pattern in MATCH_PATTERNS.items():
        if pattern.match(base_name):
            return tab_name
    return None


def extract_date_from_filename(filename):
    """Pulls a DDMMYY or DDMMYYYY date out of the trailing digits of a filename,
    e.g. bc290726 / bc_290726 / bc_29072026 / PE_300726.csv -> a date object.
    Used to pick the LATEST file automatically when two dated copies of the
    same tab (e.g. bc300726 and bc290726) are uploaded together.
    Returns None if no parseable trailing date is found."""
    base = filename
    for ext in ('.csv.gz', '.csv.zip', '.csv', '.xlsx', '.xls'):
        if base.lower().endswith(ext):
            base = base[: -len(ext)]
            break
    match = re.search(r'(\d{6,8})$', base)
    if not match:
        return None
    digits = match.group(1)
    if len(digits) not in (6, 8):
        return None
    try:
        day = int(digits[0:2])
        month = int(digits[2:4])
        year_digits = digits[4:]
        year = int(year_digits)
        if len(year_digits) == 2:
            year += 2000
        return datetime(year, month, day).date()
    except ValueError:
        return None


def auto_detect_format(col_name, df_sample_series):
    """Categorizes metadata patterns to determine cell styling formats."""
    name_clean = str(col_name).lower().strip()

    if name_clean == 'band':
        return 'number'
    # IMPORTANT: check 'date' keywords BEFORE 'price' keywords. Columns like
    # "52_Week_High_Date" or "52_Week_Low_DT" contain "high"/"low" and would
    # otherwise be misclassified as price columns, which is exactly what was
    # causing the CM_52_wk_High_low date columns to stay as raw 2-digit-year
    # text instead of being converted to proper dates.
    if any(k in name_clean for k in ['date', 'dt', 'time']):
        return 'date'
    elif any(k in name_clean for k in ['price', 'close', 'open', 'high', 'low', 'vwap', 'prev', 'trdval', 'hi_52', 'lo_52']):
        return 'price'
    elif any(k in name_clean for k in ['qty', 'volume', 'shares', 'traded', 'count', 'trades']):
        return 'qty'
    elif any(k in name_clean for k in ['percent', 'pct', 'chg%', 'return', 'yield']):
        return 'percent'
    elif any(k in name_clean for k in ['ratio', 'pe', 'pb', 'dividend']):
        return 'ratio'
    elif 'crore' in name_clean:
        return 'crores'
    elif 'lakh' in name_clean:
        return 'lakhs'
    elif pd.api.types.is_numeric_dtype(df_sample_series):
        return 'number'
    return 'text'


# =====================================================================================
# 3. Generic tabular-file loader
#    Many NSE source files ship with 1-2 "disclaimer" / "effective date" lines above
#    the real header row. Instead of hard-coding this to a single tab
#    (as the previous version did for Eligible_T0_Securities only), we now DETECT
#    the real header row for every file by trying increasing skiprows counts until
#    pandas can parse it as a proper multi-column table.
# =====================================================================================
def load_tabular_file(file_bytes, filename, max_preamble_lines=8):
    """Returns (dataframe, preamble_lines) where preamble_lines is a list of the
    raw text lines that appeared above the real header (or None if there weren't any)."""
    is_csv = filename.lower().endswith('.csv')

    def _read(skip):
        if is_csv:
            return pd.read_csv(io.BytesIO(file_bytes), skiprows=skip)
        else:
            return pd.read_excel(io.BytesIO(file_bytes), skiprows=skip)

    last_err = None
    for skip in range(0, max_preamble_lines + 1):
        try:
            df = _read(skip)
        except Exception as e:
            last_err = e
            continue
        # A real header row should give us more than 1 usable column.
        if df.shape[1] > 1:
            preamble_lines = None
            if skip > 0:
                if is_csv:
                    text = file_bytes.decode('utf-8', errors='replace')
                    preamble_lines = text.splitlines()[:skip]
                else:
                    # For Excel, just remember the raw first-column cell(s) as a header note
                    raw = pd.read_excel(io.BytesIO(file_bytes), nrows=skip, header=None)
                    preamble_lines = [str(x) for x in raw.iloc[:, 0].tolist()]
            return df, preamble_lines
        # single-column parse succeeded but isn't useful yet -> keep trying more skips

    if last_err:
        raise last_err
    raise ValueError(f"Could not detect a valid header row in {filename}")


def parse_start_cell(text):
    """Parses an Excel-style cell reference like 'B9' into (row_number, col_index).
    row_number is 1-indexed (as typed). col_index is 0-indexed (A=0, B=1, ...).
    Returns (None, None) if the text is blank or not a valid cell reference —
    in that case the tab loads normally starting from A1."""
    if not text or not str(text).strip():
        return None, None
    match = re.match(r'^\s*([A-Za-z]+)\s*(\d+)\s*$', str(text).strip())
    if not match:
        return None, None
    col_letters, row_str = match.group(1).upper(), match.group(2)
    row_number = int(row_str)
    col_index = 0
    for ch in col_letters:
        col_index = col_index * 26 + (ord(ch) - ord('A') + 1)
    col_index -= 1  # convert to 0-indexed
    if row_number < 1 or col_index < 0:
        return None, None
    return row_number, col_index


def load_tabular_file_from_cell(file_bytes, filename, row_number, col_index):
    """Loads a file starting from an explicit cell (e.g. B9): rows above it and
    columns to its left are discarded entirely — they never reach the final sheet.
    The given row becomes the header row."""
    is_csv = filename.lower().endswith('.csv')
    skip = row_number - 1
    if is_csv:
        df = pd.read_csv(io.BytesIO(file_bytes), skiprows=skip)
    else:
        df = pd.read_excel(io.BytesIO(file_bytes), skiprows=skip)
    if col_index > 0:
        df = df.iloc[:, col_index:]
    return df


def parse_row_selector(text, n_rows):
    """Parses a string like '2,5,10-15' — 1-indexed row numbers as shown in the
    on-screen preview table — into a set of 0-indexed row positions to drop.
    Invalid tokens are ignored; out-of-range numbers are clipped silently."""
    indices = set()
    if not text:
        return indices
    for part in re.split(r'[,\s]+', text.strip()):
        if not part:
            continue
        if '-' in part:
            bounds = part.split('-', 1)
            try:
                start, end = int(bounds[0]), int(bounds[1])
            except ValueError:
                continue
            if start > end:
                start, end = end, start
            for i in range(start, end + 1):
                idx0 = i - 1
                if 0 <= idx0 < n_rows:
                    indices.add(idx0)
        else:
            try:
                i = int(part)
            except ValueError:
                continue
            idx0 = i - 1
            if 0 <= idx0 < n_rows:
                indices.add(idx0)
    return indices


def clean_dataframe(df):
    """Removes fully-empty rows and fully-empty columns."""
    df = df.dropna(axis=0, how='all')
    df = df.dropna(axis=1, how='all')
    # also drop columns that are entirely empty strings / whitespace
    def _all_blank(col):
        return col.astype(str).str.strip().replace({'nan': ''}).eq('').all()
    blank_cols = [c for c in df.columns if _all_blank(df[c])]
    if blank_cols:
        df = df.drop(columns=blank_cols)
    return df.reset_index(drop=True)


# =====================================================================================
# 4. Zip / gzip ingestion
#    Wraps extracted bytes in an object that looks enough like a Streamlit
#    UploadedFile (has .name and .getvalue()) so the rest of the pipeline
#    doesn't need to know whether a file came from a direct upload or a zip.
# =====================================================================================
class InMemoryFile:
    def __init__(self, name, data: bytes):
        self.name = name
        self._data = data

    def getvalue(self):
        return self._data


def extract_all_files(zip_bytes, _depth=0, _max_depth=3):
    """Recursively walks a zip archive, expanding any nested .zip or .gz members,
    and returns a flat list of InMemoryFile objects for every leaf file found."""
    results = []
    if _depth > _max_depth:
        return results
    try:
        with zipfile.ZipFile(io.BytesIO(zip_bytes)) as zf:
            for info in zf.infolist():
                if info.is_dir():
                    continue
                name = os.path.basename(info.filename)
                if not name:
                    continue
                data = zf.read(info.filename)
                lower = name.lower()
                if lower.endswith('.zip'):
                    results.extend(extract_all_files(data, _depth + 1, _max_depth))
                elif lower.endswith('.gz'):
                    try:
                        decompressed = gzip.decompress(data)
                        results.append(InMemoryFile(name[:-3], decompressed))
                    except Exception:
                        pass  # skip unreadable gzip members
                else:
                    results.append(InMemoryFile(name, data))
    except zipfile.BadZipFile:
        pass
    return results


# =====================================================================================
# --- UI Setup Layer ---
# =====================================================================================
col_up1, col_up2 = st.columns(2)
with col_up1:
    uploaded_files = st.file_uploader(
        "Upload Source CSV or Excel Files",
        accept_multiple_files=True,
        type=['csv', 'xlsx', 'xls']
    )
with col_up2:
    uploaded_zips = st.file_uploader(
        "…or upload a ZIP / folder-export containing the source files",
        accept_multiple_files=True,
        type=['zip']
    )

all_candidate_files = list(uploaded_files) if uploaded_files else []

if uploaded_zips:
    for z in uploaded_zips:
        extracted = extract_all_files(z.getvalue())
        all_candidate_files.extend(extracted)
    st.caption(
        f"📦 Extracted {sum(len(extract_all_files(z.getvalue())) for z in uploaded_zips)} "
        f"file(s) from {len(uploaded_zips)} zip archive(s), including nested zip/gz files."
    )

# Active File Verification Status Dashboard
st.subheader("📋 Sequence Checklist & Missing Files Audit")
status_cols = st.columns(3)

# Group every uploaded/extracted file by the tab it matches. A tab can have
# more than one candidate (e.g. bc300726 AND bc290726 uploaded together) —
# we keep all of them so the person can choose, and rank them so the LATEST
# date (parsed from the trailing digits in the filename) comes first.
tab_candidates = {}
for f in all_candidate_files:
    matched_tab = resolve_tab_name(f.name)
    if matched_tab:
        tab_candidates.setdefault(matched_tab, []).append(f)


def _candidate_sort_key(f):
    d = extract_date_from_filename(f.name)
    if d is None:
        return (1, 0)  # no parseable date -> sink to the bottom, keep upload order
    return (0, -d.toordinal())  # dated files: latest date first


for files in tab_candidates.values():
    files.sort(key=_candidate_sort_key)

valid_files_map = {}

for idx, tab in enumerate(TAB_SEQUENCE):
    col_to_use = status_cols[idx % 3]
    if tab in tab_candidates:
        candidates = tab_candidates[tab]
        if len(candidates) > 1:
            labels = []
            for f in candidates:
                d = extract_date_from_filename(f.name)
                date_tag = f" — {d.strftime('%d-%b-%Y')}" if d else ""
                labels.append(f"{f.name}{date_tag}")
            sel_idx = col_to_use.selectbox(
                f"⚠️ {len(candidates)} files found for {tab} — latest chosen by default:",
                options=list(range(len(candidates))),
                format_func=lambda i, _labels=labels: _labels[i],
                key=f"select_{tab}",
            )
            chosen_file = candidates[sel_idx]
        else:
            chosen_file = candidates[0]
        valid_files_map[tab] = chosen_file
        col_to_use.markdown(f"**✅ {tab}** <small style='color:green;'>({chosen_file.name})</small>", unsafe_allow_html=True)
    else:
        col_to_use.markdown(f"**❌ {tab}** — <span style='color:#d9534f; font-weight:bold;'>Missing</span>", unsafe_allow_html=True)

st.markdown("---")

# =====================================================================================
# 5. Main Tab — quick jump navigation to every section, plus a URL reference box
# =====================================================================================
st.subheader("🏠 Main Tab — Quick Navigation")
st.caption(
    "Click a section below to jump straight to it. Every section further down has an "
    "'⬆️ Back to Main Tab' button to return here."
)

nav_cols = st.columns(3)
for idx, tab in enumerate(TAB_SEQUENCE):
    nav_col = nav_cols[idx % 3]
    with nav_col:
        if tab in valid_files_map:
            if st.button(f"➡️ {tab}", key=f"jump_{tab}", use_container_width=True):
                jump_to(f"tab_{tab}")
        else:
            st.button(
                f"🚫 {tab}",
                key=f"jump_disabled_{tab}",
                use_container_width=True,
                disabled=True,
                help="Not available — upload a matching file first.",
            )

st.markdown("---")

# --- URL reference list box (editable) ---
st.subheader("🔗 NSE Reference URLs")
st.caption("Default NSE links are pre-loaded below. Add your own URLs with the box underneath.")

DEFAULT_URLS = [
    "https://www.nseindia.com/market-data/stocks-traded",
    "https://www.nseindia.com/all-reports/",
    "https://www.nseindia.com/static/market-data/securities-available-for-trading",
    "https://www.nseindia.com/market-data/live-equity-market",
    "https://www.nseindia.com/market-data/pre-open-market-cm-and-emerge-market",
]

if "custom_urls" not in st.session_state:
    st.session_state["custom_urls"] = DEFAULT_URLS.copy()

url_in_col, url_btn_col = st.columns([5, 1])
with url_in_col:
    new_url = st.text_input(
        "Add a URL",
        key="new_url_input",
        placeholder="https://...",
        label_visibility="collapsed",
    )
with url_btn_col:
    if st.button("➕ Add URL", key="add_url_btn", use_container_width=True):
        candidate = (new_url or "").strip()
        if candidate and candidate not in st.session_state["custom_urls"]:
            st.session_state["custom_urls"].append(candidate)
            st.rerun()

for i, url in enumerate(st.session_state["custom_urls"]):
    row_link_col, row_del_col = st.columns([9, 1])
    with row_link_col:
        st.markdown(f"🔗 [{url}]({url})")
    with row_del_col:
        if st.button("🗑️", key=f"del_url_{i}", help="Remove this URL"):
            st.session_state["custom_urls"].pop(i)
            st.rerun()

st.markdown("---")

if all_candidate_files:
    if not valid_files_map:
        st.warning("⚠️ Bypassed all uploaded files. None of the file names match the required target criteria.")
    else:
        columns_to_remove = {}
        processed_dataframes = {}
        special_headers = {}  # tab -> list[str] preamble lines detected above the header

        st.subheader("🛠️ Component Tuning & Data Previews")

        for tab in TAB_SEQUENCE:
            if tab in valid_files_map:
                f = valid_files_map[tab]

                # "Zero Zero" start-cell override — some tabs default to a specific
                # start cell as a special case; every other tab defaults to blank
                # (normal A1 / "zero zero" start) unless the person sets one.
                st.session_state.setdefault(f"start_cell_{tab}", DEFAULT_START_CELLS.get(tab, ""))
                start_cell_value = st.session_state.get(f"start_cell_{tab}", "")
                start_row_number, start_col_index = parse_start_cell(start_cell_value)

                try:
                    file_bytes = f.getvalue()
                    if start_row_number is not None:
                        # Manual crop: rows above and columns left of this cell are dropped entirely.
                        df = load_tabular_file_from_cell(file_bytes, f.name, start_row_number, start_col_index)
                    else:
                        df, preamble_lines = load_tabular_file(file_bytes, f.name)
                        if preamble_lines:
                            special_headers[tab] = preamble_lines
                except Exception as e:
                    st.error(f"Failed to cleanly digest workspace segment {tab}: {e}")
                    continue

                # Strip stray whitespace from column names (e.g. sec_bhavdata_full's " SERIES")
                df.columns = [str(c).strip() for c in df.columns]

                # Remove fully-empty rows/columns per request
                df = clean_dataframe(df)

                st.markdown(f'<div id="tab_{tab}"></div>', unsafe_allow_html=True)
                section_expanded = st.session_state.get("scroll_target") == f"tab_{tab}"
                with st.expander(f"Sheet Setup Profile: {tab} ({f.name})", expanded=section_expanded):
                    if st.button("⬆️ Back to Main Tab", key=f"back_{tab}"):
                        jump_to("main_tab")

                    st.text_input(
                        f"Start cell for {tab} (optional — e.g. B9 means data starts at column B, "
                        "row 9; everything above row 9 and left of column B is dropped and won't "
                        "appear in the final file). Leave blank for a normal A1 start.",
                        key=f"start_cell_{tab}"
                    )

                    cols = df.columns.tolist()

                    col_ctrl, row_ctrl = st.columns(2)

                    with col_ctrl:
                        selected_removals = st.multiselect(
                            f"Columns to remove from {tab}:",
                            options=cols,
                            default=default_removals_for(tab, cols),
                            key=f"remove_{tab}"
                        )
                    columns_to_remove[tab] = selected_removals
                    df_after_cols = df.drop(columns=selected_removals)

                    with row_ctrl:
                        row_text = st.text_input(
                            "Rows to remove (row #s shown in the preview, e.g. 2,5,10-15):",
                            key=f"rowsel_{tab}"
                        )

                    dedupe_col, filter_col_ui, filter_val_ui = st.columns([1, 1, 2])
                    with dedupe_col:
                        remove_dupes = st.checkbox("Remove duplicate rows", key=f"dedupe_{tab}")
                    with filter_col_ui:
                        filter_col = st.selectbox(
                            "Remove rows matching a value in:",
                            options=["(none)"] + cols,
                            key=f"filtercol_{tab}"
                        )
                    exclude_values = []
                    if filter_col != "(none)":
                        unique_vals = sorted(
                            df_after_cols[filter_col].dropna().astype(str).unique().tolist()
                        )[:500]
                        with filter_val_ui:
                            exclude_values = st.multiselect(
                                f"Value(s) to remove from '{filter_col}':",
                                options=unique_vals,
                                key=f"filtervals_{tab}"
                            )

                    df_cleaned = df_after_cols.copy()
                    if remove_dupes:
                        df_cleaned = df_cleaned.drop_duplicates()
                    if exclude_values:
                        df_cleaned = df_cleaned[~df_cleaned[filter_col].astype(str).isin(exclude_values)]
                    df_cleaned = df_cleaned.reset_index(drop=True)

                    rows_to_drop = parse_row_selector(row_text, len(df_cleaned))
                    if rows_to_drop:
                        df_cleaned = df_cleaned.drop(index=sorted(rows_to_drop)).reset_index(drop=True)

                    seq_order = render_column_sequencer(
                        f"colorder_{tab}",
                        df_cleaned.columns.tolist(),
                        allow_delete=False,
                        label=f"Column order for {tab}",
                        default_order=default_order_for(tab, df_cleaned.columns.tolist()),
                    )
                    df_cleaned = df_cleaned[seq_order]

                    processed_dataframes[tab] = df_cleaned
                    st.caption(f"Rows: {len(df)} original → {len(df_cleaned)} after cleanup")
                    st.dataframe(df_cleaned.head(10), use_container_width=True)

        # -----------------------------------------------------------------------
        # Optional AI analysis (uses the Anthropic API; needs an API key)
        # -----------------------------------------------------------------------
        st.markdown("---")
        st.subheader("🤖 AI Analysis (optional)")
        st.caption(
            "Generates a short natural-language summary of each tab (row counts, notable "
            "highs/lows, anomalies). Requires an Anthropic API key — get one at "
            "console.anthropic.com. The key is only kept in this browser session; it is "
            "never written to disk or committed to GitHub."
        )
        api_key = st.text_input("Anthropic API key", type="password", key="anthropic_api_key")
        run_ai = st.button("Generate AI Insights")

        if run_ai:
            if not api_key:
                st.warning("Please enter an Anthropic API key first.")
            else:
                try:
                    import anthropic
                except ImportError:
                    st.error(
                        "The `anthropic` package isn't installed. Run `pip install anthropic` "
                        "in this app's environment and reload the page."
                    )
                else:
                    client = anthropic.Anthropic(api_key=api_key)
                    for tab, df_target in processed_dataframes.items():
                        stats_snippet = df_target.describe(include='all').to_string()[:3000]
                        with st.spinner(f"Analyzing {tab}…"):
                            try:
                                resp = client.messages.create(
                                    model="claude-sonnet-4-6",
                                    max_tokens=400,
                                    messages=[{
                                        "role": "user",
                                        "content": (
                                            f"Here are summary statistics for a financial data tab named "
                                            f"'{tab}' with {len(df_target)} rows and columns "
                                            f"{list(df_target.columns)}:\n\n{stats_snippet}\n\n"
                                            "In 3-4 short bullet points, call out anything notable "
                                            "(unusual spreads, missing data, outliers)."
                                        )
                                    }]
                                )
                                text = "".join(
                                    b.text for b in resp.content if getattr(b, "type", "") == "text"
                                )
                                with st.expander(f"AI insights — {tab}"):
                                    st.markdown(text)
                            except Exception as e:
                                st.error(f"AI analysis failed for {tab}: {e}")

        st.markdown("---")
        st.subheader("🔀 Master_Dashboard-8 — column order & inclusion")
        st.caption(
            "Master_Dashboard-8's columns come from a fixed field list, not from your "
            "uploads, so you can sequence them any time. Move columns left/right to "
            "change their order in the final sheet, or delete ones you don't want. "
            "'Symbol' is the join key and can't be deleted."
        )
        master_col_order = render_column_sequencer(
            "master_col_order",
            [f["label"] for f in MASTER_FIELD_MAP],
            allow_delete=True,
            protected=["Symbol"],
            label="Master_Dashboard-8 columns",
        )

        st.markdown("---")

        if st.button("🚀 Execute Structural Consolidation", type="primary"):
            output_stream = io.BytesIO()

            NAV_ROW = 1  # excel row 1 on every data sheet is reserved for the "⬆️ Main Tab" jump-back link
            LINK_FONT = Font(color="0563C1", underline="single", bold=True)

            with pd.ExcelWriter(output_stream, engine='openpyxl') as writer:
                exportable_tabs = [t for t in TAB_SEQUENCE if t in processed_dataframes]

                # --- Build the "Main Tab" hub sheet first, so it opens as sheet #1 ---
                main_ws = writer.book.create_sheet(title="Main Tab")
                main_ws["A1"] = "📊 Master Financial Data — Main Tab"
                main_ws["A1"].font = Font(bold=True, size=14)
                main_ws["A3"] = "Click a tab name below to jump straight to that sheet."
                main_ws["A3"].font = Font(italic=True)

                header_row_num = 5
                fixed_headers = ["Tab Name", "Rows", "Columns"]
                max_col_count = max(
                    (len(processed_dataframes[t].columns) for t in exportable_tabs), default=0
                )
                column_headers = [f"Column {i}" for i in range(1, max_col_count + 1)]
                all_headers = fixed_headers + column_headers

                for c_idx, label in enumerate(all_headers, start=1):
                    main_ws.cell(row=header_row_num, column=c_idx, value=label).font = Font(bold=True)

                for i, tab in enumerate(exportable_tabs, start=1):
                    row = header_row_num + i
                    tab_df = processed_dataframes[tab]

                    link_cell = main_ws.cell(row=row, column=1, value=f"➡️ {tab}")
                    link_cell.hyperlink = f"#'{tab}'!A1"
                    link_cell.font = LINK_FONT
                    main_ws.cell(row=row, column=2, value=len(tab_df))
                    main_ws.cell(row=row, column=3, value=len(tab_df.columns))

                    for c_idx, col_name in enumerate(tab_df.columns, start=4):
                        main_ws.cell(row=row, column=c_idx, value=str(col_name))

                main_ws.column_dimensions['A'].width = 45
                main_ws.column_dimensions['B'].width = 12
                main_ws.column_dimensions['C'].width = 12
                for c_idx in range(4, 4 + max_col_count):
                    main_ws.column_dimensions[main_ws.cell(row=header_row_num, column=c_idx).column_letter].width = 22

                for tab in TAB_SEQUENCE:
                    if tab in processed_dataframes:
                        df_target = processed_dataframes[tab]

                        preamble_offset = len(special_headers.get(tab, []))
                        start_row = NAV_ROW + preamble_offset  # 0-indexed pandas startrow

                        if tab in special_headers:
                            workbook = writer.book
                            worksheet = workbook.create_sheet(title=tab)
                            writer.sheets[tab] = worksheet
                            for i, line in enumerate(special_headers[tab], start=1):
                                worksheet.cell(row=NAV_ROW + i, column=1, value=line)

                        df_target.to_excel(writer, sheet_name=tab, startrow=start_row, index=False)

                        worksheet = writer.sheets[tab]

                        # Row 1: jump-back link to the Main Tab hub sheet
                        nav_cell = worksheet.cell(row=NAV_ROW, column=1, value="⬆️ Main Tab")
                        nav_cell.hyperlink = "#'Main Tab'!A1"
                        nav_cell.font = LINK_FONT

                        header_row = start_row + 1
                        data_start_row = header_row + 1

                        # Loop cells individually to fix numbers and scrub formatting flaws
                        for col_idx, col_name in enumerate(df_target.columns, start=1):
                            fmt_type = auto_detect_format(col_name, df_target[col_name])
                            excel_format = NUMBER_FORMATS.get(fmt_type, '@')

                            for row_idx in range(data_start_row, worksheet.max_row + 1):
                                cell = worksheet.cell(row=row_idx, column=col_idx)
                                val = cell.value

                                if val is None:
                                    continue

                                val_clean = str(val).strip() if isinstance(val, str) else val

                                if str(col_name).strip().lower() == 'band':
                                    try:
                                        cell.value = int(float(val_clean))
                                        cell.number_format = '0'
                                    except (ValueError, TypeError):
                                        cell.value = val_clean
                                        cell.number_format = '@'

                                elif fmt_type == 'date':
                                    if val_clean in ['-', '', 'NA']:
                                        cell.value = val_clean
                                        cell.number_format = '@'
                                    else:
                                        try:
                                            date_obj = pd.to_datetime(val_clean, dayfirst=True).to_pydatetime()
                                            cell.value = date_obj
                                            cell.number_format = excel_format
                                        except Exception:
                                            cell.value = val_clean
                                            cell.number_format = '@'

                                elif fmt_type in ['price', 'qty', 'percent', 'ratio', 'crores', 'lakhs', 'number']:
                                    try:
                                        numeric_clean = (
                                            str(val_clean).replace(',', '')
                                            if isinstance(val_clean, str) else val_clean
                                        )
                                        cell.value = float(numeric_clean)
                                        cell.number_format = excel_format
                                    except (ValueError, TypeError):
                                        cell.value = val_clean
                                        cell.number_format = '@'

                                else:
                                    try:
                                        numeric_clean = (
                                            str(val_clean).replace(',', '')
                                            if isinstance(val_clean, str) else val_clean
                                        )
                                        cell.value = float(numeric_clean)
                                        cell.number_format = NUMBER_FORMATS['number']
                                    except (ValueError, TypeError):
                                        cell.value = val_clean
                                        cell.number_format = excel_format

                        # Auto-adjust column widths
                        for col_idx, _ in enumerate(df_target.columns, start=1):
                            column_letter = worksheet.cell(row=header_row, column=col_idx).column_letter
                            max_length = 0
                            for row_cell in worksheet[column_letter]:
                                try:
                                    if len(str(row_cell.value)) > max_length:
                                        max_length = len(str(row_cell.value))
                                except Exception:
                                    pass
                            worksheet.column_dimensions[column_letter].width = min(max_length + 2, 50)

                        # Freeze the header row and add an auto-filter across the header
                        worksheet.freeze_panes = worksheet.cell(row=data_start_row, column=1).coordinate
                        last_col_letter = worksheet.cell(row=header_row, column=len(df_target.columns)).column_letter
                        worksheet.auto_filter.ref = f"A{header_row}:{last_col_letter}{worksheet.max_row}"

                        # Default cell this tab opens/scrolls to in Excel (A1 unless overridden above).
                        view_cell = get_default_view_cell(tab)
                        worksheet.sheet_view.topLeftCell = view_cell
                        if worksheet.sheet_view.selection:
                            worksheet.sheet_view.selection[0].activeCell = view_cell
                            worksheet.sheet_view.selection[0].sqref = view_cell

                # -----------------------------------------------------------------
                # Auto-build Master_Dashboard-8 by default — no extra click needed.
                # Reads straight off writer.book, which already holds every tab
                # just written above, and appends the joined sheet to it.
                # -----------------------------------------------------------------
                master_df, master_log = md_build_master_dashboard(writer.book)
                active_master_order = st.session_state.get(
                    "master_col_order", [f["label"] for f in MASTER_FIELD_MAP]
                )
                md_write_master_sheet(writer.book, master_df, column_order=active_master_order)

            st.success("✅ Consolidation and Formatting Complete!")

            if master_log:
                with st.expander(f"⚠️ Master_Dashboard-8: {len(master_log)} warning(s)"):
                    for line in master_log:
                        st.write("- " + line)

            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Master_Dashboard-8 Symbols", len(master_df))
            m2.metric("Columns", len(active_master_order))
            if "Symbol" in master_df.columns:
                m3.metric("Duplicate Symbols", int(master_df["Symbol"].duplicated().sum()))
                m4.metric("Blank Symbols", int((master_df["Symbol"].astype(str).str.strip() == "").sum()))
            st.dataframe(master_df[active_master_order].head(20), use_container_width=True)

            st.download_button(
                label="📥 Download Formatted Master File",
                data=output_stream.getvalue(),
                file_name=f"Master_Financial_Data_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )

# =====================================================================================
# 6. Scroll-to-anchor execution — runs last so every anchor already exists in the DOM.
#    Consumes and clears the pending scroll target so it only fires once per click.
# =====================================================================================
_scroll_target = st.session_state.get("scroll_target")
if _scroll_target:
    st.session_state["scroll_target"] = None
    components.html(
        f"""
        <script>
            setTimeout(function() {{
                var el = window.parent.document.getElementById("{_scroll_target}");
                if (el) {{ el.scrollIntoView({{behavior: "smooth", block: "start"}}); }}
            }}, 150);
        </script>
        """,
        height=0,
    )
