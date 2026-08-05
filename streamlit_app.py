import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime, date
from pathlib import Path
import uuid
import re

st.set_page_config(page_title="BUGHASHVILI Brew Journal", layout="wide")

# Bump on every deploy. Shown in the sidebar so it is obvious at a glance
# whether Streamlit Cloud is serving the latest build or a stale one.
APP_VERSION = "2026-08-04 · v12 (სისწრაფე: batch + ფორმები)"

# ============================================================
# GLOBAL DESIGN SYSTEM — one CSS block, loaded once for the whole app.
# Colors/typography come from .streamlit/config.toml; this only covers
# what config can't express. Selectors are stable: data-testid (Streamlit
# public test hooks) and .st-key-* (from widget key=), NOT version-specific
# emotion hashes.
# ============================================================
@st.cache_data
def _load_css():
    # Read the stylesheet from disk. Kept OUT of this file on purpose: a long
    # triple-quoted string here makes inspect.getsource() slices explode with
    # TokenError on Python 3.11 (Streamlit Cloud) when line numbers shift.
    f = Path(__file__).parent / "assets" / "style.css"
    return f.read_text(encoding="utf-8") if f.exists() else ""

_css = _load_css()
if _css:
    st.html("<style>" + _css + "</style>")

# ============================================================
# CONNECTION — reads credentials + sheet id from Streamlit secrets
# Local: .streamlit/secrets.toml
# Cloud: App settings -> Secrets (same format, pasted in)
# ============================================================
SCOPES = ["https://www.googleapis.com/auth/spreadsheets",
          "https://www.googleapis.com/auth/drive"]

def flash(msg):
    # st.success() immediately before st.rerun() is discarded — the rerun wipes
    # it before it ever paints. Queue it instead and show it after the rerun.
    st.session_state["_flash"] = str(msg)

_flash_msg = st.session_state.pop("_flash", None)
if _flash_msg:
    st.toast(_flash_msg, icon="✅")

@st.cache_resource
def get_client():
    creds = Credentials.from_service_account_info(
        st.secrets["gcp_service_account"], scopes=SCOPES
    )
    return gspread.authorize(creds)

@st.cache_resource
def get_sheet():
    client = get_client()
    return client.open_by_key(st.secrets["sheet_id"])

@st.cache_resource(show_spinner=False)
def ws(name):
    # Get a worksheet, creating it with headers if missing (except core sheets).
    # Cached: sh.worksheet() does a full fetch_sheet_metadata() call, so an
    # uncached ws() spends one API read *per call* — with 14 sheets and several
    # get_df/append_row per render that alone exhausted the read quota (429).
    sh = get_sheet()
    try:
        return sh.worksheet(name)
    except gspread.WorksheetNotFound:
        headers = SHEET_HEADERS.get(name)
        if headers is None:
            raise
        w = sh.add_worksheet(title=name, rows=1000, cols=len(headers))
        w.append_row(headers)
        return w

@st.cache_data(ttl=15, show_spinner=False)
def _all_sheet_values():
    # ONE batched read of every worksheet. Reading sheets one by one cost ~3s
    # per rerun (a call each); batched it is ~0.9s total, and every get_df in
    # the same rerun is then served from this single response.
    sh = get_sheet()
    names = [w.title for w in sh.worksheets()]
    resp = sh.values_batch_get([f"'{n}'!A1:AZ5000" for n in names])
    out = {}
    for name, vr in zip(names, resp.get("valueRanges", [])):
        out[name] = vr.get("values", [])
    return out

@st.cache_data(ttl=300, show_spinner=False)
def sheet_headers(name):
    # Header row rarely changes; caching it saves a read on every write.
    rows = _all_sheet_values().get(name) or []
    return rows[0] if rows else ws(name).row_values(1)

@st.cache_data(ttl=15, show_spinner=False)
def get_df(name):
    rows = _all_sheet_values().get(name) or []
    if not rows:
        return pd.DataFrame()
    headers = rows[0]
    body = []
    for r in rows[1:]:
        r = list(r) + [""] * (len(headers) - len(r))   # batch_get trims blanks
        body.append(r[:len(headers)])
    if not body:
        # empty sheet: keep the real columns so downstream df["col"]
        # filters don't KeyError
        return pd.DataFrame(columns=headers)
    df = pd.DataFrame(body, columns=headers)
    # values arrive as strings; restore numeric dtypes so existing arithmetic
    # and sorting behave exactly as they did with get_all_records()
    for c in df.columns:
        conv = pd.to_numeric(df[c], errors="coerce")
        if conv.notna().all() and (df[c].astype(str).str.strip() != "").all():
            df[c] = conv
    return df

def append_row(name, row_dict):
    """Append a row, aligning values to the sheet's existing header order."""
    headers = sheet_headers(name)
    row = [row_dict.get(h, "") for h in headers]
    ws(name).append_row(row, value_input_option="USER_ENTERED")
    get_df.clear()  # invalidate cache so the new row shows immediately

def update_cell_by_key(name, key_col, key_val, target_col, value):
    w = ws(name)
    headers = sheet_headers(name)
    key_idx = headers.index(key_col) + 1
    target_idx = headers.index(target_col) + 1
    cell = w.find(str(key_val), in_column=key_idx)
    if cell:
        w.update_cell(cell.row, target_idx, value)
        get_df.clear()  # invalidate cache so the change shows immediately
        return True
    return False

SHEET_HEADERS = {
    "WATER_PROFILES": ["Profile_Name", "Stream", "Volume_L_Default", "Gypsum_g",
                        "CaCl2_g", "Lactic_ml", "Target_pH", "Notes"],
    "CHANGE_LOG": ["Timestamp", "Operator", "Brew_ID", "Brew_Name", "Brew_Day",
                   "Action", "Details"],
    # ---- B2B client module ----
    "CLIENTS": ["Client_ID", "Name", "Type", "Contact_Person", "Phone", "Address",
                "Status", "Payment_Terms_Days", "Created_Date", "Notes"],
    "PRODUCTS": ["Product_ID", "Name", "Category", "Unit", "Default_Price_GEL",
                 "Active", "Notes"],
    "CLIENT_PRICES": ["Client_ID", "Product_ID", "Price_GEL_per_L", "Updated_Date", "Notes"],
    "SHIPMENTS": ["Shipment_ID", "Date", "Client_ID", "Product_ID", "Volume_L",
                  "Price_per_L", "Total_GEL", "Paid_Now_GEL", "Kegs_Out", "Notes",
                  "Operator", "Kegs_Returned"],
    "PAYMENTS": ["Payment_ID", "Date", "Client_ID", "Amount_GEL", "Method",
                 "Shipment_ID", "Notes", "Operator"],
    "ASSET_MOVES": ["Move_ID", "Date", "Client_ID", "Asset_Type", "Detail",
                    "Direction", "Qty", "Notes", "Operator"],
}

# equipment lent to clients; the UI also allows typing a new type
ASSET_TYPES = ["კეგი", "მაცივარი", "დიდი მაცივარი", "პეგასი", "კობრა",
               "მაგიდა (სადგამი)", "CO2 ბალონი"]
PRODUCT_CATEGORIES = ["ლუდი", "ღვინო", "არაყი", "კონიაკი", "სხვა"]

def keg_size_l():
    # all kegs are the same size; configurable in SETTINGS
    return _num(get_setting("Keg_Size_L", 30), 30) or 30

def kegs_total():
    return _num(get_setting("Kegs_Total", 58), 58)

def kegs_at_clients(moves_df=None):
    # every keg that went out and hasn't come back, across all clients
    moves_df = get_df("ASSET_MOVES") if moves_df is None else moves_df
    if moves_df.empty or "Asset_Type" not in moves_df.columns:
        return 0.0
    out = 0.0
    for _, r in moves_df[moves_df["Asset_Type"] == "კეგი"].iterrows():
        q = _num(r.get("Qty"))
        out += -q if str(r.get("Direction", "")).startswith("დაბრ") else q
    return round(out, 2)

def kegs_free():
    # kegs sitting at the brewery right now
    return round(kegs_total() - kegs_at_clients(), 2)

def client_balance(client_id, ship_df=None, pay_df=None):
    # Owed = everything shipped minus everything paid. Never stored, always
    # recomputed, so the balance can't drift out of sync with the records.
    ship_df = get_df("SHIPMENTS") if ship_df is None else ship_df
    pay_df = get_df("PAYMENTS") if pay_df is None else pay_df
    billed = paid = 0.0
    if not ship_df.empty and "Client_ID" in ship_df.columns:
        s = ship_df[ship_df["Client_ID"] == client_id]
        billed = sum(_num(x) for x in s.get("Total_GEL", []))
        paid += sum(_num(x) for x in s.get("Paid_Now_GEL", []))
    if not pay_df.empty and "Client_ID" in pay_df.columns:
        p = pay_df[pay_df["Client_ID"] == client_id]
        paid += sum(_num(x) for x in p.get("Amount_GEL", []))
    return round(billed - paid, 2)

def client_assets(client_id, moves_df=None):
    # Net equipment currently held by a client: shipped out minus returned.
    moves_df = get_df("ASSET_MOVES") if moves_df is None else moves_df
    held = {}
    if moves_df.empty or "Client_ID" not in moves_df.columns:
        return held
    for _, r in moves_df[moves_df["Client_ID"] == client_id].iterrows():
        t = str(r.get("Asset_Type", "")).strip()
        if not t:
            continue
        q = _num(r.get("Qty"))
        held[t] = held.get(t, 0) + (-q if str(r.get("Direction", "")).startswith("დაბრ") else q)
    return {k: v for k, v in held.items() if abs(v) > 1e-9}

def client_price(client_id, product_id, prices_df=None, products_df=None):
    # Per-client price if one is set, otherwise the product's default.
    prices_df = get_df("CLIENT_PRICES") if prices_df is None else prices_df
    if not prices_df.empty and {"Client_ID", "Product_ID"} <= set(prices_df.columns):
        m = prices_df[(prices_df["Client_ID"] == client_id)
                      & (prices_df["Product_ID"] == product_id)]
        if not m.empty:
            return _num(m.iloc[-1].get("Price_GEL_per_L"))
    products_df = get_df("PRODUCTS") if products_df is None else products_df
    if not products_df.empty and "Product_ID" in products_df.columns:
        m = products_df[products_df["Product_ID"] == product_id]
        if not m.empty:
            return _num(m.iloc[0].get("Default_Price_GEL"))
    return 0.0

def set_client_price(client_id, product_id, price):
    # Upsert the client's current price for a product.
    w = ws("CLIENT_PRICES")
    values = w.get_all_values()
    hdr = values[0]
    ci, pi, pr = hdr.index("Client_ID"), hdr.index("Product_ID"), hdr.index("Price_GEL_per_L")
    di = hdr.index("Updated_Date")
    for ridx, rowv in enumerate(values[1:], start=2):
        if len(rowv) > max(ci, pi) and rowv[ci] == client_id and rowv[pi] == product_id:
            w.update_cell(ridx, pr + 1, price)
            w.update_cell(ridx, di + 1, str(date.today()))
            get_df.clear()
            return
    append_row("CLIENT_PRICES", {"Client_ID": client_id, "Product_ID": product_id,
                                 "Price_GEL_per_L": price, "Updated_Date": str(date.today())})

BREW_NUM_RE = re.compile(r"^ხარშვა\s+(\d+)\s*—\s*(.*)$")

def parse_brew_number(display_name):
    """Extract N from 'ხარშვა N — Style'. Returns None if it doesn't match."""
    m = BREW_NUM_RE.match(str(display_name).strip())
    return int(m.group(1)) if m else None

def next_brew_number(headers_df):
    """Highest existing brew number + 1 (survives deletions, unlike a row count)."""
    if headers_df.empty or "Display_Name" not in headers_df.columns:
        return 1
    nums = [parse_brew_number(d) for d in headers_df["Display_Name"]]
    nums = [n for n in nums if n]
    return max(nums) + 1 if nums else 1

def shift_brew_numbers_from(n):
    # Renumber every brew with number >= n up by one, freeing slot n so a
    # forgotten brew can be inserted in the middle. Returns rows changed.
    w = ws("BREW_HEADER")
    values = w.get_all_values()
    if not values:
        return 0
    hdr = values[0]
    try:
        di = hdr.index("Display_Name")
    except ValueError:
        return 0
    updates = []
    for ridx, rowv in enumerate(values[1:], start=2):
        if len(rowv) <= di:
            continue
        m = BREW_NUM_RE.match(str(rowv[di]).strip())
        if m and int(m.group(1)) >= n:
            updates.append((ridx, f"ხარშვა {int(m.group(1)) + 1} — {m.group(2)}"))
    # descending so the highest numbers move first — no transient collisions
    for ridx, newname in sorted(updates, reverse=True):
        w.update_cell(ridx, di + 1, newname)
    if updates:
        get_df.clear()
    return len(updates)

def log_change(action, details="", bid="", brew_name="", day=""):
    # Append an audit-trail row to CHANGE_LOG. Deliberately does NOT clear
    # the get_df cache — the log is never read in the UI hot path, and clearing
    # would force every other sheet to refetch on each save.
    try:
        w = ws("CHANGE_LOG")
        w.append_row([datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                      st.session_state.get("operator", ""), str(bid), str(brew_name),
                      str(day), str(action), str(details)],
                     value_input_option="USER_ENTERED")
    except Exception as e:  # logging must never block the actual save
        st.caption(f"⚠️ ჟურნალში ჩაწერა ვერ მოხერხდა: {type(e).__name__}")

def new_id(prefix):
    return f"{prefix}-{datetime.now().strftime('%Y%m%d-%H%M%S')}-{uuid.uuid4().hex[:4]}"

# plausibility limits for one 800 L brew day — implausible values are almost
# always typos (wrong unit, extra digit, wrong field)
SANITY = {
    "vol_target": 800.0, "vol_tolerance": 0.30,   # ±30% of the day's volume
    "sparge_max": 2000.0, "grain_max_kg": 300.0, "grain_min_kg": 50.0,
    "hop_max_g": 3000.0, "salt_max_g": 1000.0, "lactic_max_ml": 2000.0,
    "mash_temp_min": 30.0, "mash_temp_max": 85.0, "mash_dur_max": 180,
    "og_max_p": 30.0, "ph_min": 4.0, "ph_max": 7.0,
}

def sanity_gate(checks, key):
    # checks = list of (is_suspicious, message). Renders warnings and requires
    # an explicit confirm checkbox before allowing the save. Returns True when it
    # is safe to proceed (nothing suspicious, or the user confirmed).
    hits = [msg for bad, msg in checks if bad]
    if not hits:
        return True
    for msg in hits:
        st.warning(f"⚠️ {msg}")
    return st.checkbox("დიახ, მონაცემი სწორია — შენახვა", key=key)

def _num(v, default=0.0):
    """Coerce a sheet cell (str like '1000' or '5.40%', int, or blank) to float."""
    if v in (None, ""):
        return default
    try:
        return float(str(v).replace("%", "").replace(",", "").strip())
    except (ValueError, TypeError):
        return default

UNIT_TO_GRAMS_OR_ML = {"kg": 1000, "g": 1, "l": 1000, "ml": 1, "pack": 1}
# units are only interconvertible inside the same group; anything across
# groups (or unknown) returns None so we never deduct blindly
_UNIT_GROUP = {"kg": "mass", "g": "mass", "l": "volume", "ml": "volume", "pack": "pack"}

def convert_to_inventory_unit(qty, from_unit, inventory_unit):
    # Convert qty from the form's unit to the INVENTORY row's unit via a
    # common base (grams/milliliters). Returns None when the units are
    # incompatible (e.g. 'pack' vs 'kg') — caller must block and warn instead
    # of computing blindly.
    fu = str(from_unit).strip().lower()
    iu = str(inventory_unit).strip().lower()
    if fu not in UNIT_TO_GRAMS_OR_ML or iu not in UNIT_TO_GRAMS_OR_ML:
        return None
    if _UNIT_GROUP[fu] != _UNIT_GROUP[iu]:
        return None
    return qty * UNIT_TO_GRAMS_OR_ML[fu] / UNIT_TO_GRAMS_OR_ML[iu]

def find_inventory_item(inv_df, patterns):
    # Find an INVENTORY row whose Item name contains any of the (lowercase)
    # patterns. Returns the row (Series) or None. Name-tolerant on purpose —
    # e.g. 'GYpsum' matches 'gypsum'.
    if inv_df.empty or "Item" not in inv_df.columns:
        return None
    for _, r in inv_df.iterrows():
        name = str(r.get("Item", "")).lower()
        if any(p in name for p in patterns):
            return r
    return None

def delete_rows_where(sheet_name, col_name, value):
    # Delete every row in a worksheet whose col_name equals value.
    # Collects 1-based row indices first, then deletes bottom-up so indices
    # don't shift. Returns the number of deleted rows.
    w = ws(sheet_name)
    values = w.get_all_values()
    if not values:
        return 0
    try:
        ci = values[0].index(col_name)
    except ValueError:
        return 0
    idxs = [i + 1 for i, r in enumerate(values)
            if i > 0 and len(r) > ci and r[ci] == value]
    for i in reversed(idxs):
        w.delete_rows(i)
    if idxs:
        get_df.clear()
    return len(idxs)

def is_brew_finished(brew_id, headers_df):
    # Single source of truth for 'finished': FG_Date is set (i.e. the
    # 'ეს არის FG' checkbox was used). Used by the edit-lock gate.
    match = headers_df[headers_df["Brew_ID"] == brew_id]
    if match.empty:
        return False
    return str(match.iloc[0].get("FG_Date", "")).strip() not in ("", "nan", "None")

def lock_gate(fg_date, key):
    # Render the finished-brew warning + override checkbox for a tab.
    # Returns True when the tab should stay locked (inputs disabled).
    st.warning(f"ეს ხარშვა დასრულებულია (FG დაფიქსირდა {fg_date}-ზე). "
               f"ცვლილება არ არის რეკომენდებული.")
    return not st.checkbox("მაინც მინდა რედაქტირება", key=key)

def is_day_closed(row, day):
    """True when this brew day was explicitly closed (Day{N}_Closed set)."""
    return str(row.get(f"Day{int(day)}_Closed", "")).strip().lower() in ("1", "true", "yes", "დახურული")

def close_brew_day(bid, day, closed=True):
    """Mark a brew day closed/open in BREW_HEADER."""
    update_cell_by_key("BREW_HEADER", "Brew_ID", bid, f"Day{int(day)}_Closed",
                       "1" if closed else "")
    get_df.clear()

def delete_day_rows(sheet, bid, day, category=None):
    # Delete rows for Brew_ID + Brew_Day (optionally + Category) so a day's
    # data can be replaced on re-save. Bottom-up delete; returns count.
    w = ws(sheet)
    values = w.get_all_values()
    if not values:
        return 0
    h = values[0]
    try:
        bi, di = h.index("Brew_ID"), h.index("Brew_Day")
        ci = h.index("Category") if category is not None else None
    except ValueError:
        return 0
    idxs = [i + 1 for i, r in enumerate(values)
            if i > 0 and len(r) > max(bi, di) and r[bi] == bid and r[di] == str(day)
            and (ci is None or (len(r) > ci and r[ci] == category))]
    for i in reversed(idxs):
        w.delete_rows(i)
    if idxs:
        get_df.clear()
    return len(idxs)

def water_defaults_for_day(bid, day, fallback):
    # Return saved {Mash:{...}, Sparge:{...}} water values for a brew day.
    # If day 2 has nothing saved, fall back to day 1's saved values; if neither,
    # use `fallback` (profile / manual defaults). Enables day-2 pre-fill.
    df = get_df("WATER_TREATMENT")
    def rows_for(d):
        if df.empty or "Brew_Day" not in df.columns:
            return {}
        sub = df[(df["Brew_ID"] == bid) & (df["Brew_Day"].astype(str) == str(d))]
        out = {}
        for _, r in sub.iterrows():
            out[str(r.get("Water_Stream"))] = r
        return out
    saved = rows_for(day) or (rows_for(1) if str(day) != "1" else {})
    result = {}
    for stream in ("Mash", "Sparge"):
        r = saved.get(stream)
        if r is not None:
            result[stream] = {"vol": _num(r.get("Volume_L")), "gyp": _num(r.get("Gypsum_g")),
                              "cacl2": _num(r.get("CaCl2_g")), "lac": _num(r.get("Lactic_ml")),
                              "ph": _num(r.get("Target_pH"), fallback[stream]["ph"])}
        else:
            result[stream] = fallback[stream]
    return result

def mash_steps_df(bid, day):
    # 5-row editor DataFrame prefilled with a brew day's saved Mash_Step rows.
    # If day 2 has none, prefill from day 1 (day-2 defaults to day-1).
    df = get_df("BREW_STEPS")
    rows = []
    if not df.empty and {"Brew_ID", "Category", "Brew_Day"} <= set(df.columns):
        def sub_for(d):
            return df[(df["Brew_ID"] == bid) & (df["Category"] == "Mash_Step")
                      & (df["Brew_Day"].astype(str) == str(d))]
        sub = sub_for(day)
        if sub.empty and str(day) != "1":
            sub = sub_for(1)
        for _, r in sub.iterrows():
            rows.append({"ტემპ (°C)": _num(r.get("Qty")),
                         "ხანგრძლივობა (წთ)": int(_num(str(r.get("Timing", "")).replace("წთ", ""))),
                         "შენიშვნა": str(r.get("Justification", "") or "")})
    while len(rows) < 5:
        rows.append({"ტემპ (°C)": 0.0, "ხანგრძლივობა (წთ)": 0, "შენიშვნა": ""})
    out = pd.DataFrame(rows[:5])
    out.insert(0, "საფეხური #", [1, 2, 3, 4, 5])
    return out

def copy_ingredients_day1_to_day2(bid, category, form_unit):
    # Duplicate day-1 rows of a category (Malt/Hop) into day 2. Day 2 is a
    # real second 800 L batch, so it consumes inventory again — deducts per
    # item, skipping any that exceed stock (unless historical mode).
    # Returns (copied, skipped) item-name lists.
    steps = get_df("BREW_STEPS")
    needed = {"Brew_ID", "Category", "Brew_Day", "Item", "Qty"}
    if steps.empty or not needed <= set(steps.columns):
        return [], []
    d1 = steps[(steps["Brew_ID"] == bid) & (steps["Category"] == category)
               & (steps["Brew_Day"].astype(str) == "1")]
    inv = get_df("INVENTORY")
    hist = st.session_state.get("historical_mode", False)
    copied, skipped = [], []
    for _, r in d1.iterrows():
        item = r["Item"]
        qty = _num(r.get("Qty"))
        inv_row = (inv[inv["Item"] == item] if not inv.empty and "Item" in inv.columns
                   else pd.DataFrame())
        unit = str(inv_row.iloc[0].get("Unit", "") or "").strip() if not inv_row.empty else form_unit
        stock = _num(inv_row.iloc[0].get("Current_Qty", 0)) if not inv_row.empty else 0
        deduct = convert_to_inventory_unit(qty, form_unit, unit)
        if not hist and (deduct is None or deduct > stock):
            skipped.append(item)
            continue
        newrow = {c: r.get(c) for c in steps.columns if c}
        newrow["Brew_Day"] = 2
        append_row("BREW_STEPS", newrow)
        if not hist and not inv_row.empty:
            update_cell_by_key("INVENTORY", "Item", item, "Current_Qty",
                               round(stock - deduct, 4))
        copied.append(item)
    return copied, skipped

def update_step_qty(bid, category, item, new_qty, timing=None, brew_day=None):
    # Update the Qty of the BREW_STEPS row matching Brew_ID + Category + Item
    # (a multi-column match — update_cell_by_key can't do that, it keys on one
    # column). When timing is given (hops) the Timing column must match too, and
    # when brew_day is given the Brew_Day column must match: the same ingredient
    # on another brew day is a separate addition, not a duplicate.
    # Returns True if a row was found and updated.
    w = ws("BREW_STEPS")
    values = w.get_all_values()
    if not values:
        return False
    headers = values[0]
    try:
        bi, ci, ii, qi = (headers.index("Brew_ID"), headers.index("Category"),
                          headers.index("Item"), headers.index("Qty"))
        ti = headers.index("Timing") if timing is not None else None
        di = headers.index("Brew_Day") if brew_day is not None else None
    except ValueError:
        return False
    for ridx, rowv in enumerate(values[1:], start=2):
        if (len(rowv) > max(bi, ci, ii) and rowv[bi] == bid
                and rowv[ci] == category and rowv[ii] == item
                and (ti is None or (len(rowv) > ti and rowv[ti] == str(timing)))
                and (di is None or (len(rowv) > di and rowv[di] == str(brew_day)))):
            w.update_cell(ridx, qi + 1, new_qty)
            get_df.clear()
            return True
    return False

def get_setting(key, default=None):
    """Read a value from the SETTINGS worksheet (Setting/Value columns)."""
    df = get_df("SETTINGS")
    if df.empty:
        return default
    match = df[df["Setting"] == key]
    if match.empty:
        return default
    return match.iloc[0]["Value"]

def _hydro_density_factor(t):
    # Water density-correction polynomial used by the sheet's hydrometer
    # temperature correction (t in °C).
    return (1.00130346 - 0.000134722124 * t + 0.00000204052596 * t ** 2
            - 0.00000000232820948 * t ** 3)

def correct_gravity_plato(raw_p, temp_c, ref_temp_c):
    # Port of the BREW_GRAVITY_LOG Gravity_P_Corrected sheet formula.
    # Converts a raw °P hydrometer reading at temp_c to °P corrected to the
    # hydrometer's reference temperature. Returns rounded float, or None if
    # inputs are missing.
    if raw_p in ("", None) or temp_c in ("", None):
        return None
    d = float(raw_p)
    e = float(temp_c)
    b = float(ref_temp_c)
    sg = 1 + d / (258.6 - (d / 258.2) * 227.1)
    corrected_sg = sg * _hydro_density_factor(e) / _hydro_density_factor(b)
    plato = (-616.868 + 1111.14 * corrected_sg - 630.272 * corrected_sg ** 2
             + 135.997 * corrected_sg ** 3)
    return round(plato, 2)

def calc_header_metrics(actual_og_p, actual_fg_p, post_boil_vol_l, total_grain_kg, malt_yield_fraction):
    # Compute BREW_HEADER derived metrics in Python (formerly sheet formulas
    # that only existed on rows 2-22). Returns only the metrics whose inputs
    # are present.
    def sg(p):
        return 1 + p / (258.6 - (p / 258.2) * 227.1)
    result = {}
    if actual_og_p and actual_fg_p:
        result["ADF_%"] = round((actual_og_p - actual_fg_p) / actual_og_p * 100, 1)
        result["Est_ABV_%"] = round((sg(actual_og_p) - sg(actual_fg_p)) * 131.25, 1)
    if actual_og_p and post_boil_vol_l and total_grain_kg and malt_yield_fraction:
        result["Brewhouse_Eff_%"] = round(
            (post_boil_vol_l * sg(actual_og_p) * actual_og_p / 100) /
            (total_grain_kg * malt_yield_fraction) * 100, 2)
    return result

def total_grain_kg(bid):
    """SUMIFS of BREW_STEPS Qty where Brew_ID matches and Category == 'Malt'."""
    df = get_df("BREW_STEPS")
    if df.empty or "Category" not in df.columns:
        return 0.0
    m = df[(df["Brew_ID"] == bid) & (df["Category"] == "Malt")]
    return round(sum(_num(x) for x in m["Qty"]), 2) if not m.empty else 0.0

def recalc_header_metrics(bid, og, fg, post_boil_vol, grain_kg, fg_confirmed=False):
    # Recompute and write Total_Grain_kg + the three derived metrics into
    # BREW_HEADER for one brew. og/fg/post_boil_vol/grain_kg are coerced; missing
    # inputs simply skip the metrics they feed.
    #
    # Eff/ADF/ABV are FG-dependent and only meaningful once FG is *confirmed*
    # (the "ეს არის FG" checkbox, i.e. FG_Date is set). Pass fg_confirmed=True
    # only in that case; otherwise those three are skipped so a provisional
    # interim reading never lands in the header. Total_Grain_kg is written
    # regardless — it is valid independently of FG.
    myf = _num(get_setting("Malt_Yield_Fraction", 0.8), 0.8)
    metrics = calc_header_metrics(_num(og), _num(fg), _num(post_boil_vol), _num(grain_kg), myf)
    if grain_kg:
        update_cell_by_key("BREW_HEADER", "Brew_ID", bid, "Total_Grain_kg", grain_kg)
    if not fg_confirmed:
        return {}  # FG not approved yet — do not write FG-dependent metrics
    for col in ("ADF_%", "Est_ABV_%", "Brewhouse_Eff_%"):
        if col in metrics:
            update_cell_by_key("BREW_HEADER", "Brew_ID", bid, col, metrics[col])
    return metrics

# ============================================================
# SIDEBAR NAV
# ============================================================
page = st.sidebar.radio("გვერდი", ["📦 Inventory", "🍺 ხარშვა", "🤝 კლიენტები"],
                        key="nav_radio")

# ---- historical entry mode: record old brews without touching stock ----
st.session_state.setdefault("historical_mode", False)
st.session_state.historical_mode = st.sidebar.checkbox(
    "📜 ისტორიული ხარშვის შეყვანა (მარაგს არ ჩამოაკლებს)",
    value=st.session_state.historical_mode, key="historical_mode_cb"
)
if st.session_state.historical_mode:
    st.sidebar.warning("ისტორიული რეჟიმი ჩართულია — ხარშვა ჩაიწერება, "
                       "მაგრამ მარაგს არ ჩამოაკლდება. გამორთე, როცა ძველი "
                       "ხარშვების შეყვანას დაასრულებ.")
hist_mode = st.session_state.historical_mode

# who is entering data — recorded in CHANGE_LOG (the app has no login)
st.sidebar.text_input("👤 ვინ ავსებს (ჟურნალისთვის)", key="operator",
                      placeholder="სახელი")
st.sidebar.caption(f"ვერსია: {APP_VERSION}")

# ============================================================
# PAGE 1 — INVENTORY
# ============================================================
if page == "📦 Inventory":
    st.title("📦 მარაგი (Inventory)")
    st.caption("[Certain] აქ ცვლილება პირდაპირ INVENTORY sheet-ს სწორედება. "
               "ავტომატური ჩამოჭრა ხარშვიდან ჯერ არ არის ჩართული — განზრახ.")

    df = get_df("INVENTORY")
    if df.empty:
        st.warning("INVENTORY ცარიელია.")
    else:
        cat_filter = st.multiselect("კატეგორია", options=sorted(df["Category"].unique()),
                                     default=list(df["Category"].unique()))
        view = df[df["Category"].isin(cat_filter)] if cat_filter else df
        edited = st.data_editor(
            view, num_rows="fixed", use_container_width=True,
            disabled=["Low_Stock_Flag"], key="inv_editor"
        )
        if st.button("💾 ცვლილებების შენახვა"):
            w = ws("INVENTORY")
            headers = sheet_headers("INVENTORY")
            full = df.copy()
            full.update(edited)
            values = [headers] + full[headers].astype(str).values.tolist()
            w.clear()
            w.update(values)
            get_df.clear()
            log_change("მარაგის ცხრილი შესწორდა", f"{len(full)} row გადაიწერა")
            flash("შენახულია.")
            st.rerun()

    st.divider()
    st.subheader("➕ ახალი ნედლეულის დამატება")
    with st.form("add_item"):
        c1, c2, c3 = st.columns(3)
        item = c1.text_input("დასახელება")
        cat = c2.selectbox("კატეგორია", ["Hop", "Malt", "Salt", "Yeast", "Other"])
        unit = c3.selectbox("ერთეული", ["g", "kg", "ml", "l", "pack"])
        c4, c5, c6 = st.columns(3)
        qty = c4.number_input("რაოდენობა", min_value=0.0, step=1.0)
        low = c5.number_input("Low threshold", min_value=0.0, step=1.0)
        aa = c6.text_input("AA% (თუ ჰოპია)")
        c7, c8 = st.columns(2)
        ebc = c7.text_input("EBC (თუ ალაოა)")
        notes = c8.text_input("შენიშვნა")
        submitted = st.form_submit_button("დამატება")
        if submitted and item:
            append_row("INVENTORY", {
                "Item": item, "Category": cat, "Unit": unit, "Current_Qty": qty,
                "AA_%_if_hop": aa, "EBC": ebc, "Low_Threshold": low,
                "Received_Date": str(date.today()), "Notes": notes,
            })
            flash(f"{item} დაემატა.")
            st.rerun()

    if not df.empty:
        with st.expander("🗑️ ნედლეულის წაშლა"):
            st.warning("ეს წაშლის item-ს INVENTORY-დან. ისტორიული BREW_STEPS "
                       "ჩანაწერები უცვლელი რჩება. ქმედება შეუქცევადია.")
            item_to_delete = st.selectbox("აირჩიე", df["Item"].tolist(), key="inv_del_sel")
            confirm = st.checkbox(f"დავადასტურებ „{item_to_delete}“-ის წაშლას",
                                  key="inv_del_confirm")
            if st.button("წაშლა", key="inv_del_btn") and confirm:
                n = delete_rows_where("INVENTORY", "Item", item_to_delete)
                get_df.clear()
                log_change("🗑️ ნედლეული წაიშალა", f"{item_to_delete} ({n} row)")
                flash(f"წაშლილია ({n} row).")
                st.rerun()

# ============================================================
# PAGE 2 — BREW
# ============================================================
elif page == "🍺 ხარშვა":
    st.title("🍺 ხარშვის ჟურნალი")

    headers_df = get_df("BREW_HEADER")
    # dropdown shows Display_Name, internally maps back to Brew_ID
    label_to_bid = {}
    if not headers_df.empty:
        for _, hr in headers_df.iterrows():
            dn = str(hr.get("Display_Name", "") or "").strip()
            lab = dn if dn and dn.lower() != "nan" else str(hr["Brew_ID"])
            if lab in label_to_bid:  # guard against duplicate names
                lab = f"{lab} ({str(hr['Brew_ID'])[-4:]})"
            label_to_bid[lab] = hr["Brew_ID"]
    choice = st.selectbox("აირჩიე ხარშვა ან შექმენი ახალი",
                          ["➕ ახალი ხარშვა"] + list(label_to_bid.keys()))

    # ---------- NEW BREW ----------
    if choice == "➕ ახალი ხარშვა":
        st.subheader("ახალი ხარშვის დაწყება")
        # number lives OUTSIDE the form so the insert notice updates live
        # (form widgets don't rerun until submit)
        next_no = next_brew_number(headers_df)
        brew_no = st.number_input(
            "ხარშვის ნომერი", min_value=1, value=next_no, step=1, key="new_brew_no",
            help="დატოვე როგორც არის ბოლოში დასამატებლად. თუ დაკავებულ ნომერს "
                 "აირჩევ (გამორჩენილი ხარშვის ჩასამატებლად), ის და ყველა "
                 "მომდევნო ხარშვა ავტომატურად ერთით გადაინომრება.")
        if brew_no < next_no:
            st.info(f"↩️ ჩასმა პოზიცია {int(brew_no)}-ზე: ამჟამინდელი „ხარშვა {int(brew_no)}“ "
                    f"და ყველა მომდევნო ერთით გადაინომრება "
                    f"(→ {int(brew_no) + 1}, {int(brew_no) + 2}, …).")
        with st.form("new_brew"):
            c1, c2, c3 = st.columns(3)
            b_date = c1.date_input("თარიღი", value=date.today())
            style = c2.text_input("ლუდის სტილი", placeholder="მაგ. Märzen/Oktoberfest")
            ferm = c3.selectbox("ფერმენტორი", ["CCT1", "CCT2"])
            c4, c5 = st.columns(2)
            target_vol = c4.number_input("სამიზნე მოცულობა (L)", min_value=0.0, value=800.0)
            water_us = c5.number_input("წყლის Water_uS", min_value=0.0, value=70.0)
            c6, c7 = st.columns(2)
            target_og = c6.number_input("Target OG (°P)", min_value=0.0, step=0.1)
            target_fg = c7.number_input("Target FG (°P)", min_value=0.0, step=0.1)

            start = st.form_submit_button("ხარშვის დაწყება")
            if start and style:
                bid = new_id("BREW")
                shifted = 0
                if brew_no < next_no:  # inserting into the middle — free the slot
                    shifted = shift_brew_numbers_from(int(brew_no))
                display_name = f"ხარშვა {int(brew_no)} — {style}"
                append_row("BREW_HEADER", {
                    "Brew_ID": bid, "Date": str(b_date), "Beer_Style": style,
                    "Fermenter": ferm, "Target_Vol_L": target_vol, "Water_uS": water_us,
                    "Target_OG_P": target_og, "Target_FG_P": target_fg,
                    "Display_Name": display_name,
                })
                log_change("ხარშვა შეიქმნა",
                           f"{display_name}" + (f" (ჩასმა — {shifted} ხარშვა გადაინომრა)" if shifted else ""),
                           bid=bid, brew_name=display_name)
                flash(f"ხარშვა შეიქმნა: {display_name} ({bid})."
                           + (f" {shifted} ხარშვა გადაინომრა." if shifted else "")
                           + " აირჩიე ის ზემოთა სიიდან რომ დეტალები შეავსო.")
                st.rerun()

    # ---------- EXISTING BREW: TABS ----------
    else:
        bid = label_to_bid[choice]
        row = headers_df[headers_df["Brew_ID"] == bid].iloc[0]
        fg_done = is_brew_finished(bid, headers_df)

        st.subheader(choice)  # Display_Name already carries the style
        st.caption(f"Brew_ID: {bid}")
        if fg_done:
            st.info("✅ FG მიღწეულია — ეს ხარშვა დახურულია. ცვლილება კვლავ შესაძლებელია, მაგრამ საჭირო აღარ არის.")

        # ── single global brew-day switch in the top panel ──────────────
        # one control drives ALL per-day data (წყალი/მეშინგი/ალაო/სვია).
        # OG/Boil-header, საფუარი, gravity are shared across both days.
        with st.container(border=True, key="brewday-panel"):
            # a widget's session_state key can only be set BEFORE the widget is
            # created, so closing day 1 leaves a flag that is applied here
            if st.session_state.pop("_jump_to_day2", False):
                st.session_state["brew_day_global"] = 2
            brew_day = st.radio(
                "📅 ხარშვის დღე (თითო 800 L) — გადართე და ყველა ველი შესაბამის დღეზე გადავა",
                [1, 2], horizontal=True, key="brew_day_global")

            d1_closed, d2_closed = is_day_closed(row, 1), is_day_closed(row, 2)
            st.caption(("🔒 დღე 1 დახურულია" if d1_closed else "🔓 დღე 1 ღიაა")
                       + " · " + ("🔒 დღე 2 დახურულია" if d2_closed else "🔓 დღე 2 ღიაა"))

            this_day_closed = is_day_closed(row, brew_day)
            # day 2 stays locked until day 1 is finished and closed — you can't
            # start the second batch while the first one is still open
            day2_blocked = (brew_day == 2 and not d1_closed)
            # closed day => everything per-day is locked until explicitly reopened
            day_locked = this_day_closed or day2_blocked
            if this_day_closed:
                st.warning(f"დღე {brew_day} დახურულია — შემთხვევითი ცვლილებისგან დაცულია. "
                           f"რედაქტირებისთვის მონიშნე ქვემოთ.")
                if st.checkbox(f"🔓 მაინც მინდა დღე {brew_day}-ის რედაქტირება",
                               key=f"day_override_{bid}_{brew_day}"):
                    day_locked = False
                    if st.button(f"დღე {brew_day}-ის საბოლოოდ გახსნა", key=f"reopen_{bid}_{brew_day}"):
                        close_brew_day(bid, brew_day, closed=False)
                        log_change("დღე გაიხსნა", f"დღე {brew_day}", bid=bid,
                                   brew_name=choice, day=brew_day)
                        flash(f"დღე {brew_day} გაიხსნა.")
                        st.rerun()
            elif day2_blocked:
                st.warning("დღე 2 ჯერ დაბლოკილია — ჯერ დაასრულე და დახურე დღე 1.")
                if st.button("🔒 დღე 1-ის დახურვა ახლავე", key=f"close_d1_from_d2_{bid}"):
                    close_brew_day(bid, 1, closed=True)
                    log_change("დღე დაიხურა", "დღე 1 (დღე 2-ზე გადასვლისას)",
                               bid=bid, brew_name=choice, day=1)
                    flash("დღე 1 დაიხურა — დღე 2 გაიხსნა.")
                    st.rerun()
                if st.checkbox("🔓 მაინც მინდა დღე 2-ის შევსება (დღე 1-ის დახურვის გარეშე)",
                               key=f"day2_override_{bid}"):
                    day_locked = False
            else:
                if brew_day == 2:
                    st.caption("დღე 2 — ველები default-ად დღე 1-ის მონაცემებით ივსება; "
                               "შეასწორე რაც განსხვავდება. ალაო/სვიისთვის იხ. „დღე 1-ის კოპირება“ ღილაკი.")
                if st.button(f"🔒 დღე {brew_day}-ის დახურვა", key=f"close_{bid}_{brew_day}"):
                    close_brew_day(bid, brew_day, closed=True)
                    log_change("დღე დაიხურა", f"დღე {brew_day}", bid=bid,
                               brew_name=choice, day=brew_day)
                    if brew_day == 1:  # jump straight to day 2 — the step that got forgotten
                        st.session_state["_jump_to_day2"] = True
                    st.rerun()

        BREW_TABS = ["📊 მიმოხილვა", "💧 წყალი", "🌾 მეშინგი", "🔥 დუღილი (boil/hop)", "🧪 ფერმენტაცია/Gravity"]
        if "brew_tab" not in st.session_state:
            st.session_state.brew_tab = "📊 მიმოხილვა"
        st.session_state.brew_tab = st.radio(
            "ტაბი", BREW_TABS,
            horizontal=True, label_visibility="collapsed", key="brew_tab_radio",
            index=BREW_TABS.index(st.session_state.brew_tab)
        )

        # ---- OVERVIEW TAB ----
        if st.session_state.brew_tab == "📊 მიმოხილვა":
            fg_confirmed = str(row.get("FG_Date", "")).strip() not in ("", "nan", "None")

            def _show(v):
                s = str(v).strip()
                return s if s not in ("", "nan", "None") else "—"

            # per-sheet reads once, filtered locally below
            steps_all = get_df("BREW_STEPS")
            steps_b = (steps_all[steps_all["Brew_ID"] == bid]
                       if not steps_all.empty and "Brew_ID" in steps_all.columns
                       else pd.DataFrame())
            water_all = get_df("WATER_TREATMENT")
            water_b = (water_all[water_all["Brew_ID"] == bid]
                       if not water_all.empty and "Brew_ID" in water_all.columns
                       else pd.DataFrame())
            grav_all = get_df("BREW_GRAVITY_LOG")
            grav_b = (grav_all[grav_all["Brew_ID"] == bid]
                      if not grav_all.empty and "Brew_ID" in grav_all.columns
                      else pd.DataFrame())

            # --- header card ---
            with st.container(border=True, key="ovcard-header"):
                st.markdown(f"### 📊 {_show(row.get('Beer_Style'))}")
                c1, c2, c3 = st.columns(3)
                c1.markdown(f"**სტილი**  \n{_show(row.get('Beer_Style'))}")
                c2.markdown(f"**თარიღი**  \n{_show(row.get('Date'))}")
                c3.markdown(f"**ფერმენტორი**  \n{_show(row.get('Fermenter'))}")

                yeast_name_v = _show(row.get("Yeast"))
                if yeast_name_v != "—":
                    yform_v = _show(row.get("Yeast_Form"))
                    ygen_v = _show(row.get("Yeast_Generation"))
                    yline = yeast_name_v + (f" · {yform_v}" if yform_v != "—" else "")
                    if str(yform_v).startswith("ლექი") and ygen_v != "—":
                        yline += f" · თაობა {ygen_v}"
                    st.markdown(f"**🧫 საფუარი:** {yline}")

                c1, c2, c3, c4 = st.columns(4, gap="large")
                c1.metric("Target OG (°P)", _show(row.get("Target_OG_P")))
                c2.metric("Actual OG (°P)", _show(row.get("Actual_OG_P")))
                c3.metric("Target FG (°P)", _show(row.get("Target_FG_P")))
                c4.metric("Actual FG (°P)", _show(row.get("Actual_FG_P")))

                if fg_confirmed:
                    c1, c2, c3 = st.columns(3, gap="large")
                    c1.metric("Brewhouse Eff %", _show(row.get("Brewhouse_Eff_%")))
                    c2.metric("ADF %", _show(row.get("ADF_%")))
                    c3.metric("Est ABV %", _show(row.get("Est_ABV_%")))
                else:
                    st.warning("⏳ მიმდინარეობს — FG დაუდასტურებელია. "
                               "Brewhouse_Eff_%/ADF_%/Est_ABV_% ჯერ არ ითვლება.")

            # --- water summary card ---
            with st.container(border=True, key="ovcard-water"):
                st.markdown("#### 💧 წყალი")
                st.caption("ორივე დღის ჯამი (მოცულობა/მარილები); pH — საშუალო.")
                def _wsum(col):
                    return sum(_num(x) for x in water_b[col]) if (not water_b.empty and col in water_b.columns) else 0
                def _wavg(col):
                    vals = [_num(x) for x in water_b[col]] if (not water_b.empty and col in water_b.columns) else []
                    vals = [v for v in vals if v]
                    return round(sum(vals) / len(vals), 2) if vals else None
                c1, c2, c3, c4 = st.columns(4, gap="large")
                c1.metric("სულ მოცულობა (L)", f"{_wsum('Volume_L'):g}" if _wsum('Volume_L') else "—")
                c2.metric("Total Gypsum (g)", f"{_wsum('Gypsum_g'):g}" if _wsum('Gypsum_g') else "—")
                c3.metric("Total CaCl₂ (g)", f"{_wsum('CaCl2_g'):g}" if _wsum('CaCl2_g') else "—")
                c4.metric("Total Lactic (ml)", f"{_wsum('Lactic_ml'):g}" if _wsum('Lactic_ml') else "—")
                avg_ph = _wavg("Target_pH")
                st.metric("საშ. Target pH", f"{avg_ph:g}" if avg_ph else "—")
                if not water_b.empty:
                    wcols = [c for c in ["Brew_Day", "Water_Stream", "Volume_L", "Gypsum_g",
                                         "CaCl2_g", "Lactic_ml", "Target_pH"] if c in water_b.columns]
                    st.dataframe(water_b.sort_values([c for c in ["Brew_Day", "Water_Stream"]
                                 if c in water_b.columns])[wcols],
                                 use_container_width=True, hide_index=True)
                else:
                    st.caption("წყლის მონაცემი ჯერ არ არის შენახული.")

            # --- grain bill card ---
            with st.container(border=True, key="ovcard-grain"):
                st.markdown("#### 🌾 ალაო (grain bill)")
                malt_b = (steps_b[steps_b["Category"] == "Malt"]
                          if not steps_b.empty and "Category" in steps_b.columns else pd.DataFrame())
                if not malt_b.empty:
                    mcols = [c for c in ["Brew_Day", "Item", "Qty", "Unit", "Justification"]
                             if c in malt_b.columns]
                    st.dataframe(malt_b[mcols], use_container_width=True, hide_index=True)
                    total_grain = sum(_num(x) for x in malt_b["Qty"])
                    st.metric("ჯამური Total_Grain_kg", f"{total_grain:g} kg")
                else:
                    st.caption("ალაო ჯერ არ არის დამატებული.")

            # --- mash schedule card ---
            with st.container(border=True, key="ovcard-mash"):
                st.markdown("#### 🌾 Mash schedule")
                mash_b = (steps_b[steps_b["Category"] == "Mash_Step"]
                          if not steps_b.empty and "Category" in steps_b.columns else pd.DataFrame())
                if not mash_b.empty:
                    mscols = [c for c in ["Brew_Day", "Item", "Qty", "Unit", "Timing", "Justification"]
                              if c in mash_b.columns]
                    st.dataframe(mash_b[mscols], use_container_width=True, hide_index=True)
                else:
                    st.caption("Mash schedule ჯერ არ არის შევსებული.")

            # --- hop schedule card (ordered 60წთ → 0 → Whirlpool) ---
            with st.container(border=True, key="ovcard-hop"):
                st.markdown("#### 🔥 Hop schedule")
                hop_b = (steps_b[steps_b["Category"] == "Hop"]
                         if not steps_b.empty and "Category" in steps_b.columns else pd.DataFrame())
                if not hop_b.empty:
                    hop_order = {"60წთ": 0, "30წთ": 1, "15წთ": 2, "5წთ": 3, "0": 4, "Whirlpool": 5}
                    hop_b = hop_b.copy()
                    hop_b["_o"] = hop_b["Timing"].map(lambda t: hop_order.get(str(t), 99))
                    hop_b = hop_b.sort_values("_o")
                    hcols = [c for c in ["Brew_Day", "Item", "Qty", "Unit", "Timing", "AA_%",
                                         "Justification"] if c in hop_b.columns]
                    st.dataframe(hop_b[hcols], use_container_width=True, hide_index=True)
                else:
                    st.caption("ჰოპი ჯერ არ არის დამატებული.")

            # --- gravity curve card ---
            with st.container(border=True, key="ovcard-gravity"):
                st.markdown("#### 🧪 Gravity curve")
                if not grav_b.empty:
                    gv = grav_b.copy()
                    gv["Day_#"] = pd.to_numeric(gv.get("Day_#"), errors="coerce")
                    gv["Gravity_P_Corrected"] = pd.to_numeric(gv.get("Gravity_P_Corrected"), errors="coerce")
                    gv = gv.sort_values("Day_#")
                    chart = gv.dropna(subset=["Day_#", "Gravity_P_Corrected"])
                    if len(chart) >= 2:
                        st.line_chart(chart.set_index("Day_#")[["Gravity_P_Corrected"]])
                    elif len(chart) == 1:
                        st.caption("მხოლოდ ერთი გაზომვაა — მრუდისთვის საჭიროა ≥2.")
                    last_cols = [c for c in ["Day_#", "Date", "Gravity_P_Raw", "Gravity_P_Corrected",
                                             "Temp_C"] if c in gv.columns]
                    st.write("**ბოლო 3 გაზომვა:**")
                    st.dataframe(gv[last_cols].tail(3), use_container_width=True, hide_index=True)
                else:
                    st.caption("Gravity log ჯერ ცარიელია.")

            # --- outcome note card ---
            if _show(row.get("Outcome_Note")) != "—":
                with st.container(border=True, key="ovcard-outcome"):
                    st.markdown("#### 📝 შედეგის შენიშვნა")
                    st.write(_show(row.get("Outcome_Note")))

            # --- change log for this brew ---
            with st.expander("🧾 ცვლილებების ჟურნალი"):
                log_df = get_df("CHANGE_LOG")
                if log_df.empty or "Brew_ID" not in log_df.columns:
                    st.caption("ჟურნალი ჯერ ცარიელია.")
                else:
                    mine = log_df[log_df["Brew_ID"] == bid]
                    if mine.empty:
                        st.caption("ამ ხარშვაზე ჩანაწერი ჯერ არ არის.")
                    else:
                        lcols = [c for c in ["Timestamp", "Operator", "Brew_Day", "Action", "Details"]
                                 if c in mine.columns]
                        st.dataframe(mine[lcols].iloc[::-1], use_container_width=True,
                                     hide_index=True)

            # --- delete brew (cascade, double confirmation) ---
            with st.expander("🗑️ ხარშვის წაშლა"):
                st.warning("ეს წაშლის ამ ხარშვას ყველა Sheet-იდან — Header, Steps, "
                           "Water, Gravity. ეს ქმედება შეუქცევადია.")
                brew_display = str(row.get("Display_Name", "") or "").strip() or choice
                confirm_text = st.text_input(
                    f"დასადასტურებლად ჩაწერე ხარშვის სახელი: „{brew_display}“",
                    key=f"del_confirm_{bid}")
                if st.button("წაშლა საბოლოოდ", key=f"del_btn_{bid}"):
                    if confirm_text.strip() != brew_display:
                        st.error("სახელი არ ემთხვევა — წაშლა არ შესრულდა.")
                    else:
                        deleted = {}
                        for sheet in ["BREW_STEPS", "WATER_TREATMENT",
                                      "BREW_GRAVITY_LOG", "BREW_HEADER"]:
                            deleted[sheet] = delete_rows_where(sheet, "Brew_ID", bid)
                        get_df.clear()
                        log_change("🗑️ ხარშვა წაიშალა",
                                   ", ".join(f"{s} −{n}" for s, n in deleted.items()),
                                   bid=bid, brew_name=brew_display)
                        flash("წაშლილია: " + ", ".join(
                            f"{s} −{n}" for s, n in deleted.items()))
                        st.rerun()

        # ---- WATER TAB ----
        if st.session_state.brew_tab == "💧 წყალი":
            locked = (lock_gate(row.get("FG_Date"), f"edit_ovr_water_{bid}") if fg_done else False) or day_locked
            water_day = brew_day
            profiles_df = get_df("WATER_PROFILES")
            profile_names = sorted(profiles_df["Profile_Name"].unique()) if not profiles_df.empty else []
            chosen_profile = st.selectbox("წყლის პროფილი (თუ შენახული გაქვს)",
                                           ["— ხელით შევსება —"] + profile_names, disabled=locked)

            defaults = {"Mash": {"vol": 0, "gyp": 0, "cacl2": 0, "lac": 0, "ph": 5.3},
                        "Sparge": {"vol": 1000, "gyp": 0, "cacl2": 0, "lac": 0, "ph": 5.7}}
            if chosen_profile != "— ხელით შევსება —":
                for stream in ["Mash", "Sparge"]:
                    prow = profiles_df[(profiles_df["Profile_Name"] == chosen_profile) &
                                        (profiles_df["Stream"] == stream)]
                    if not prow.empty:
                        pr = prow.iloc[0]
                        defaults[stream] = {
                            "vol": pr.get("Volume_L_Default") or defaults[stream]["vol"],
                            "gyp": pr.get("Gypsum_g", 0), "cacl2": pr.get("CaCl2_g", 0),
                            "lac": pr.get("Lactic_ml", 0), "ph": pr.get("Target_pH", defaults[stream]["ph"]),
                        }
            # per-day pre-fill: saved day-N values, else day-1 values, else profile
            wd = water_defaults_for_day(bid, water_day, defaults)

            with st.form(f"water_form_{water_day}"):
                st.markdown(f"**Mash წყალი** — დღე {water_day}")
                m1, m2, m3, m4, m5 = st.columns(5)
                m_vol = m1.number_input("მოცულობა (L)", min_value=0.0, value=float(wd["Mash"]["vol"] or 0), disabled=locked)
                m_gyp = m2.number_input("Gypsum (g)", value=float(wd["Mash"]["gyp"] or 0), disabled=locked)
                m_cacl = m3.number_input("CaCl2 (g)", value=float(wd["Mash"]["cacl2"] or 0), disabled=locked)
                m_lac = m4.number_input("Lactic (ml)", value=float(wd["Mash"]["lac"] or 0), disabled=locked)
                m_ph = m5.number_input("Target pH", value=float(wd["Mash"]["ph"] or 5.3), disabled=locked)

                st.markdown(f"**Sparge წყალი** — დღე {water_day}")
                s1, s2, s3, s4, s5 = st.columns(5)
                s_vol = s1.number_input("მოცულობა (L)", value=float(wd["Sparge"]["vol"] or 1000), disabled=locked)
                s_gyp = s2.number_input("Gypsum (g)", value=float(wd["Sparge"]["gyp"] or 0), disabled=locked)
                s_cacl = s3.number_input("CaCl2 (g)", value=float(wd["Sparge"]["cacl2"] or 0), disabled=locked)
                s_lac = s4.number_input("Lactic (ml)", value=float(wd["Sparge"]["lac"] or 0), disabled=locked)
                s_ph = s5.number_input("Target pH", value=float(wd["Sparge"]["ph"] or 5.7), disabled=locked)
                water_confirm = st.checkbox("დიახ, უჩვეულო მონაცემიც სწორია", disabled=locked)
                water_submit = st.form_submit_button(
                    f"💾 წყლის მონაცემის შენახვა (დღე {water_day})", disabled=locked)

            tgt = _num(row.get("Target_Vol_L"), SANITY["vol_target"]) or SANITY["vol_target"]
            lo, hi = tgt * (1 - SANITY["vol_tolerance"]), tgt * (1 + SANITY["vol_tolerance"])
            water_problems = [m for bad, m in [
                (m_vol and not (lo <= m_vol <= hi),
                 f"Mash მოცულობა {m_vol:g} L — სამიზნეს ({tgt:g} L) 30%-ზე მეტით სცდება."),
                (s_vol > SANITY["sparge_max"],
                 f"Sparge მოცულობა {s_vol:g} L — ჩვეულებრივზე ბევრად მეტია."),
                (m_gyp + s_gyp > SANITY["salt_max_g"],
                 f"Gypsum სულ {m_gyp + s_gyp:g} g — ჩვეულებრივზე ბევრად მეტია."),
                (m_cacl + s_cacl > SANITY["salt_max_g"],
                 f"CaCl₂ სულ {m_cacl + s_cacl:g} g — ჩვეულებრივზე ბევრად მეტია."),
                (m_lac + s_lac > SANITY["lactic_max_ml"],
                 f"Lactic სულ {m_lac + s_lac:g} ml — ჩვეულებრივზე ბევრად მეტია."),
                (m_ph and not (SANITY["ph_min"] <= m_ph <= SANITY["ph_max"]),
                 f"Mash pH {m_ph:g} — რეალურ დიაპაზონს ({SANITY['ph_min']}–{SANITY['ph_max']}) სცდება."),
                (s_ph and not (SANITY["ph_min"] <= s_ph <= SANITY["ph_max"]),
                 f"Sparge pH {s_ph:g} — რეალურ დიაპაზონს სცდება."),
            ] if bad]
            if water_submit and water_problems and not water_confirm:
                for m in water_problems:
                    st.warning(f"⚠️ {m}")
                st.error("დაადასტურე ფორმაში („უჩვეულო მონაცემიც სწორია“) და ხელახლა შეინახე.")
            elif water_submit:
                # salt auto-deduction: total need across Mash+Sparge, in grams.
                # Only NEW saves deduct — existing WATER_TREATMENT rows are
                # never touched retroactively.
                # replace-on-save: this brew day's water is rewritten, not
                # duplicated. Salts deduct only on the day's FIRST save (no
                # prior rows) so editing + re-saving doesn't double-deduct.
                wt_df = get_df("WATER_TREATMENT")
                day_had_rows = (not wt_df.empty and "Brew_Day" in wt_df.columns
                                and not wt_df[(wt_df["Brew_ID"] == bid)
                                              & (wt_df["Brew_Day"].astype(str) == str(water_day))].empty)
                need = {"Gypsum": m_gyp + s_gyp, "CaCl2": m_cacl + s_cacl}
                blocked = False
                deductions = {}  # salt -> (item_name, stock, deduct_in_inv_unit, inv_unit)
                if not hist_mode and not day_had_rows:  # deduct once, on first save
                    inv_w_df = get_df("INVENTORY")
                    salt_rows = {
                        "Gypsum": find_inventory_item(inv_w_df, ["gypsum", "caso4"]),
                        "CaCl2": find_inventory_item(inv_w_df, ["cacl", "calcium chloride"]),
                    }
                    for salt, needed_g in need.items():
                        if needed_g <= 0:
                            continue
                        r = salt_rows[salt]
                        if r is None:
                            st.warning(f"INVENTORY-ში ვერ მოიძებნა {salt} (Salt) — "
                                       f"ამ მარილზე ჩამოჭრა გამოტოვებულია.")
                            continue
                        inv_unit = str(r.get("Unit", "") or "").strip()
                        stock = _num(r.get("Current_Qty", 0))
                        deduct = convert_to_inventory_unit(needed_g, "g", inv_unit)
                        if deduct is None:
                            st.error(f"{r['Item']}: ერთეულები შეუთავსებელია (g ↔ "
                                     f"'{inv_unit}') — გაასწორე INVENTORY-ში. შენახვა დაიბლოკა.")
                            blocked = True
                        elif deduct > stock:
                            st.error(f"{r['Item']} მარაგშია მხოლოდ {stock:g}{inv_unit}, "
                                     f"საჭიროა {deduct:g}{inv_unit}.")
                            blocked = True
                        else:
                            deductions[salt] = (r["Item"], stock, deduct, inv_unit)
                if not blocked:
                    if day_had_rows:  # replace: drop this day's old rows first
                        delete_day_rows("WATER_TREATMENT", bid, water_day)
                    append_row("WATER_TREATMENT", {"Brew_ID": bid, "Water_Stream": "Mash",
                        "Volume_L": m_vol, "Gypsum_g": m_gyp, "CaCl2_g": m_cacl,
                        "Lactic_ml": m_lac, "Target_pH": m_ph, "Brew_Day": water_day})
                    append_row("WATER_TREATMENT", {"Brew_ID": bid, "Water_Stream": "Sparge",
                        "Volume_L": s_vol, "Gypsum_g": s_gyp, "CaCl2_g": s_cacl,
                        "Lactic_ml": s_lac, "Target_pH": s_ph, "Brew_Day": water_day})
                    for salt, (item, stock, deduct, inv_unit) in deductions.items():
                        update_cell_by_key("INVENTORY", "Item", item,
                                           "Current_Qty", round(stock - deduct, 4))
                    log_change("წყალი შენახვა",
                               f"Mash {m_vol:g}L/{m_gyp:g}g gyp/{m_cacl:g}g cacl2/{m_lac:g}ml/pH{m_ph:g}; "
                               f"Sparge {s_vol:g}L/{s_gyp:g}g gyp/{s_cacl:g}g cacl2/{s_lac:g}ml/pH{s_ph:g}"
                               + (" [replace]" if day_had_rows else ""),
                               bid=bid, brew_name=choice, day=water_day)
                    msg = f"წყლის მონაცემი შენახულია (დღე {water_day})."
                    if hist_mode:
                        msg += " 📜 ისტორიული რეჟიმი — მარაგი უცვლელია."
                    elif day_had_rows:
                        msg += " (განახლდა — მარაგი ხელახლა არ შეხებია)."
                    elif deductions:
                        parts = ", ".join(f"{salt} {d:g}{u}"
                                          for salt, (_, _, d, u) in deductions.items())
                        msg += f" მარაგიდან ჩამოეჭრა: {parts}."
                    st.success(msg)

            st.caption("იმ პროფილს ხედავ პირველად? — შეინახე ახლანდელი მნიშვნელობები, რომ მომდევნო ჯერზე აღარ გჭირდეს ხელით.")
            new_profile_name = st.text_input("შეინახე ეს, როგორც ახალი პროფილი (სახელი)", disabled=locked)
            if st.button("💾 პროფილად შენახვა", disabled=locked) and new_profile_name:
                append_row("WATER_PROFILES", {"Profile_Name": new_profile_name, "Stream": "Mash",
                    "Volume_L_Default": m_vol, "Gypsum_g": m_gyp, "CaCl2_g": m_cacl,
                    "Lactic_ml": m_lac, "Target_pH": m_ph})
                append_row("WATER_PROFILES", {"Profile_Name": new_profile_name, "Stream": "Sparge",
                    "Volume_L_Default": s_vol, "Gypsum_g": s_gyp, "CaCl2_g": s_cacl,
                    "Lactic_ml": s_lac, "Target_pH": s_ph})
                flash(f"პროფილი '{new_profile_name}' შენახულია.")
                st.rerun()

            existing_water = get_df("WATER_TREATMENT")
            if not existing_water.empty:
                st.dataframe(existing_water[existing_water["Brew_ID"] == bid],
                             use_container_width=True)

        # ---- MASH TAB ----
        if st.session_state.brew_tab == "🌾 მეშინგი":
            locked = (lock_gate(row.get("FG_Date"), f"edit_ovr_mash_{bid}") if fg_done else False) or day_locked
            st.caption("[Certain] mash schedule ხარშვიდან ხარშვამდე იცვლება — ამიტომ ყოველთვის ხელით.")
            st.markdown(f"**საფეხურები — დღე {brew_day}** _(დღე 2 default-ად დღე 1-ის "
                        f"საფეხურებით ივსება; შენახვა ცვლის ამ დღის საფეხურებს)_")
            mash_edit = st.data_editor(
                mash_steps_df(bid, brew_day), num_rows="fixed", use_container_width=True,
                key=f"mash_steps_editor_{brew_day}", disabled=locked,
            )
            _mt = [_num(v) for v in mash_edit["ტემპ (°C)"] if _num(v)]
            _md = [_num(v) for v in mash_edit["ხანგრძლივობა (წთ)"] if _num(v)]
            mash_ok = sanity_gate([
                (any(not (SANITY["mash_temp_min"] <= t <= SANITY["mash_temp_max"]) for t in _mt),
                 f"ტემპერატურა დიაპაზონს ({SANITY['mash_temp_min']:g}–{SANITY['mash_temp_max']:g}°C) "
                 f"სცდება: {', '.join(f'{t:g}' for t in _mt if not (SANITY['mash_temp_min'] <= t <= SANITY['mash_temp_max']))}°C"),
                (any(d > SANITY["mash_dur_max"] for d in _md),
                 f"ხანგრძლივობა {SANITY['mash_dur_max']} წთ-ზე მეტია: "
                 f"{', '.join(f'{d:g}' for d in _md if d > SANITY['mash_dur_max'])} წთ"),
            ], key=f"sanity_mash_{bid}_{brew_day}")
            if st.button(f"💾 საფეხურების შენახვა (დღე {brew_day})",
                         disabled=locked or not mash_ok):
                # within-editor duplicate-temperature guard (day 1 and day 2 may
                # share temps; the whole day is replaced on save)
                to_save, batch_temps, dup_err = [], {}, None
                for _, r in mash_edit.iterrows():
                    temp_v = r["ტემპ (°C)"]
                    dur_v = r["ხანგრძლივობა (წთ)"]
                    if (temp_v in (None, "", 0, 0.0)) and (dur_v in (None, "", 0)):
                        continue  # ცარიელი row — იგნორი
                    t = _num(temp_v)
                    step_no = int(r["საფეხური #"])
                    if t in batch_temps:
                        dup_err = (f"ეს ტემპერატურა ({t:g}°C) გამეორებულია — "
                                   f"საფეხური {batch_temps[t]} და {step_no}.")
                        break
                    batch_temps[t] = step_no
                    to_save.append((step_no, temp_v, dur_v, r["შენიშვნა"]))

                if dup_err:
                    st.error(dup_err)
                else:
                    delete_day_rows("BREW_STEPS", bid, brew_day, category="Mash_Step")
                    for step_no, temp_v, dur_v, note in to_save:
                        append_row("BREW_STEPS", {
                            "Brew_ID": bid, "Stage": "Mash", "Category": "Mash_Step",
                            "Item": f"საფეხური {step_no}",
                            "Qty": temp_v, "Unit": "°C",
                            "Timing": f"{int(dur_v)} წთ", "Justification": note,
                            "Brew_Day": brew_day,
                        })
                    log_change("მეშინგის საფეხურები",
                               "; ".join(f"{_num(t):g}°C/{int(_num(d))}წთ" for _, t, d, _n in to_save)
                               or "(ცარიელი)",
                               bid=bid, brew_name=choice, day=brew_day)
                    flash(f"დღე {brew_day}: {len(to_save)} საფეხური შენახულია.")
                    st.rerun()

            st.markdown("**ალაოს ჩამონათვალი (grain bill)**")
            st.caption("[Certain] დამატებისას რაოდენობა ავტომატურად ჩამოეჭრება "
                       "INVENTORY-ს (ერთეულის კონვერტაციით). მარაგზე მეტს ვერ აირჩევ.")
            inv_df = get_df("INVENTORY")
            malt_opts = (sorted(inv_df[inv_df["Category"] == "Malt"]["Item"].tolist())
                         if not inv_df.empty and "Category" in inv_df.columns else [])
            if not malt_opts:
                st.info("INVENTORY-ში Malt კატეგორიის ჩანაწერი არ არის — ჯერ დაამატე მარაგში.")
            else:
                malt_name = st.selectbox("ალაო", malt_opts, key="malt_sel", disabled=locked)
                malt_inv = inv_df[inv_df["Item"] == malt_name].iloc[0]
                malt_stock = _num(malt_inv.get("Current_Qty", 0))
                malt_inv_unit = str(malt_inv.get("Unit", "") or "").strip()
                # stock converted into the form's unit (kg) → live cap
                malt_cap = convert_to_inventory_unit(malt_stock, malt_inv_unit, "kg")
                if malt_cap is None and not hist_mode:
                    st.error(f"ერთეულები შეუთავსებელია: ფორმაში kg, მარაგში "
                             f"'{malt_inv_unit}' — ჩამოჭრა ვერ გამოითვლება. "
                             f"გაასწორე ამ item-ის ერთეული INVENTORY-ში.")
                malt_day = brew_day
                with st.form(f"malt_form_{brew_day}", clear_on_submit=True):
                    c1, c3 = st.columns([2, 2])
                    malt_qty = c1.number_input(
                        f"რაოდენობა (kg) — მარაგში {malt_cap:g} kg" if malt_cap is not None
                        else "რაოდენობა (kg)",
                        min_value=0.0,
                        # historical entries may exceed today's stock — no cap then
                        max_value=malt_cap if (malt_cap and malt_cap > 0 and not hist_mode) else None,
                        disabled=locked)
                    malt_just = c3.text_input("დასაბუთება", disabled=locked)
                    malt_confirm = st.checkbox("დიახ, უჩვეულო რაოდენობაც სწორია", disabled=locked)
                    malt_submit = st.form_submit_button("დამატება (ალაო)", disabled=locked)
                malt_deduct = convert_to_inventory_unit(malt_qty, "kg", malt_inv_unit)
                # duplicate = same Brew_ID+Category+Item AND same brew day —
                # day 1 and day 2 of the same malt stay separate rows
                malt_steps = get_df("BREW_STEPS")
                malt_existing = (malt_steps[(malt_steps["Brew_ID"] == bid)
                                            & (malt_steps["Category"] == "Malt")
                                            & (malt_steps["Item"] == malt_name)
                                            & (malt_steps["Brew_Day"].astype(str) == str(malt_day))]
                                 if not malt_steps.empty
                                 and {"Brew_ID", "Category", "Item", "Brew_Day"} <= set(malt_steps.columns)
                                 else pd.DataFrame())
                if hist_mode:
                    st.caption("📜 ისტორიული რეჟიმი — მარაგიდან არ ჩამოეჭრება.")
                elif malt_deduct:
                    if not malt_existing.empty:
                        prev = _num(malt_existing.iloc[0].get("Qty"))
                        st.caption(f"დღე {malt_day}-ზე უკვე დამატებულია {prev:g} kg — "
                                   f"დაემატება ჯამში {prev + malt_qty:g} kg. "
                                   f"მარაგიდან ჩამოეჭრება: {malt_deduct:g} {malt_inv_unit}")
                    else:
                        st.caption(f"დღე {malt_day} · მარაგიდან ჩამოეჭრება: "
                                   f"{malt_deduct:g} {malt_inv_unit}")
                malt_odd = malt_qty > SANITY["grain_max_kg"]
                if malt_submit and malt_qty <= 0:
                    st.error("შეავსე რაოდენობა.")
                elif malt_submit and malt_odd and not malt_confirm:
                    st.warning(f"⚠️ ალაო {malt_qty:g} kg — ერთ დღეზე "
                               f"{SANITY['grain_max_kg']:g} kg-ზე მეტია. ერთეული ხომ არ აგერია (g ↔ kg)?")
                    st.error("დაადასტურე ფორმაში და ხელახლა შეინახე.")
                elif malt_submit:
                    if not hist_mode and malt_deduct is None:
                        st.error("ერთეულები შეუთავსებელია — ჩანაწერი არ შენახულა, "
                                 "ჩამოჭრა არ მომხდარა.")
                    elif not hist_mode and malt_deduct > malt_stock:
                        st.error(f"მარაგშია მხოლოდ {malt_stock:g} {malt_inv_unit}, "
                                 f"მოთხოვნილია {malt_deduct:g} {malt_inv_unit}.")
                    else:
                        if not malt_existing.empty:
                            new_total = round(_num(malt_existing.iloc[0].get("Qty")) + malt_qty, 4)
                            update_step_qty(bid, "Malt", malt_name, new_total,
                                            brew_day=malt_day)
                        else:
                            append_row("BREW_STEPS", {
                                "Brew_ID": bid, "Stage": "Mash", "Category": "Malt",
                                "Item": malt_name, "Qty": malt_qty, "Unit": "kg",
                                "Justification": malt_just, "Brew_Day": malt_day,
                            })
                        if not hist_mode:
                            update_cell_by_key("INVENTORY", "Item", malt_name,
                                               "Current_Qty", round(malt_stock - malt_deduct, 4))
                        log_change("ალაო დამატება",
                                   f"{malt_name} +{malt_qty:g} kg"
                                   + ("" if hist_mode else f" (მარაგი −{malt_deduct:g} {malt_inv_unit})"),
                                   bid=bid, brew_name=choice, day=malt_day)
                        st.rerun()

                if brew_day == 2 and not locked and st.button("📋 დღე 1-ის ალაო → დღე 2"):
                    copied, skipped = copy_ingredients_day1_to_day2(bid, "Malt", "kg")
                    if copied:
                        st.success(f"დღე 2-ზე დაკოპირდა: {', '.join(copied)} "
                                   "(მარაგიდან ჩამოიჭრა).")
                    if skipped:
                        st.warning(f"მარაგი არ ჰყოფნის, გამოტოვდა: {', '.join(skipped)}.")
                    if copied:
                        st.rerun()

            steps_df = get_df("BREW_STEPS")
            if not steps_df.empty:
                mash_view = steps_df[(steps_df["Brew_ID"] == bid) & (steps_df["Stage"] == "Mash")]
                st.dataframe(mash_view, use_container_width=True)

        # ---- BOIL / HOP TAB ----
        if st.session_state.brew_tab == "🔥 დუღილი (boil/hop)":
            # day lock covers the hop list (per-day); the boil header below is
            # brew-level and stays governed by the FG lock only
            fg_locked = lock_gate(row.get("FG_Date"), f"edit_ovr_boil_{bid}") if fg_done else False
            locked = fg_locked or day_locked
            st.markdown("**ჰოპის დამატება**")
            st.caption("[Certain] დამატებისას რაოდენობა ავტომატურად ჩამოეჭრება "
                       "INVENTORY-ს (ერთეულის კონვერტაციით). AA% ივსება INVENTORY-დან.")
            inv_hop_df = get_df("INVENTORY")
            hop_opts = (sorted(inv_hop_df[inv_hop_df["Category"] == "Hop"]["Item"].tolist())
                        if not inv_hop_df.empty and "Category" in inv_hop_df.columns else [])
            if not hop_opts:
                st.info("INVENTORY-ში Hop კატეგორიის ჩანაწერი არ არის — ჯერ დაამატე მარაგში.")
            else:
                hop_name = st.selectbox("ჰოპი", hop_opts, key="hop_sel", disabled=locked)
                hop_inv_row = inv_hop_df[inv_hop_df["Item"] == hop_name].iloc[0]
                hop_stock = _num(hop_inv_row.get("Current_Qty", 0))
                hop_inv_unit = str(hop_inv_row.get("Unit", "") or "").strip()
                hop_aa = str(hop_inv_row.get("AA_%_if_hop", "") or "")
                # stock converted into the form's unit (g) → live cap
                hop_cap = convert_to_inventory_unit(hop_stock, hop_inv_unit, "g")
                if hop_cap is None and not hist_mode:
                    st.error(f"ერთეულები შეუთავსებელია: ფორმაში g, მარაგში "
                             f"'{hop_inv_unit}' — ჩამოჭრა ვერ გამოითვლება. "
                             f"გაასწორე ამ item-ის ერთეული INVENTORY-ში.")
                with st.form(f"hop_form_{brew_day}", clear_on_submit=True):
                    c1, c2, c3 = st.columns(3)
                    hop_qty = c1.number_input(
                        f"რაოდენობა (g) — მარაგში {hop_cap:g} g" if hop_cap is not None
                        else "რაოდენობა (g)",
                        min_value=0.0,
                        # historical entries may exceed today's stock — no cap then
                        max_value=hop_cap if (hop_cap and hop_cap > 0 and not hist_mode) else None,
                        disabled=locked)
                    hop_timing_sel = c2.selectbox(
                        "დრო", ["60წთ", "30წთ", "15წთ", "5წთ", "0", "Whirlpool", "სხვა"],
                        disabled=locked)
                    c3.text_input("AA% (INVENTORY-დან)", value=hop_aa, disabled=True,
                                  key="hop_aa_display")
                    hop_timing_custom = st.text_input(
                        "დრო — ხელით (თუ „სხვა“ აირჩიე, მაგ. 45წთ)", disabled=locked)
                    hop_just = st.text_input("დასაბუთება", disabled=locked)
                    hop_confirm = st.checkbox("დიახ, უჩვეულო რაოდენობაც სწორია", disabled=locked)
                    hop_submit = st.form_submit_button("დამატება (ჰოპი)", disabled=locked)

                hop_timing = (hop_timing_custom.strip() if hop_timing_sel == "სხვა"
                              else hop_timing_sel)
                hop_day = brew_day
                hop_deduct = convert_to_inventory_unit(hop_qty, "g", hop_inv_unit)
                # true duplicate for hops = same Brew_ID+Category+Item AND Timing
                # AND brew day. Same hop at another timing OR another day is a
                # separate addition (Tradition@30წთ day1 vs day2 = two rows).
                hop_steps = get_df("BREW_STEPS")
                hop_existing = (hop_steps[(hop_steps["Brew_ID"] == bid)
                                          & (hop_steps["Category"] == "Hop")
                                          & (hop_steps["Item"] == hop_name)
                                          & (hop_steps["Timing"].astype(str) == str(hop_timing))
                                          & (hop_steps["Brew_Day"].astype(str) == str(hop_day))]
                                if not hop_steps.empty
                                and {"Brew_ID", "Category", "Item", "Timing", "Brew_Day"} <= set(hop_steps.columns)
                                else pd.DataFrame())
                if hist_mode:
                    st.caption("📜 ისტორიული რეჟიმი — მარაგიდან არ ჩამოეჭრება.")
                elif hop_deduct:
                    if not hop_existing.empty:
                        prev = _num(hop_existing.iloc[0].get("Qty"))
                        st.caption(f"დღე {hop_day} · {hop_timing}-ზე უკვე დამატებულია "
                                   f"{prev:g} g — დაემატება ჯამში {prev + hop_qty:g} g. "
                                   f"მარაგიდან ჩამოეჭრება: {hop_deduct:g} {hop_inv_unit}")
                    else:
                        st.caption(f"დღე {hop_day} · მარაგიდან ჩამოეჭრება: "
                                   f"{hop_deduct:g} {hop_inv_unit}")
                hop_odd = hop_qty > SANITY["hop_max_g"]
                if hop_submit and hop_qty <= 0:
                    st.error("შეავსე რაოდენობა.")
                elif hop_submit and hop_odd and not hop_confirm:
                    st.warning(f"⚠️ სვია {hop_qty:g} g — ერთ დამატებაზე "
                               f"{SANITY['hop_max_g']:g} g-ზე მეტია. ერთეული ხომ არ აგერია (g ↔ kg)?")
                    st.error("დაადასტურე ფორმაში და ხელახლა შეინახე.")
                elif hop_submit:
                    if not str(hop_timing).strip():
                        st.error("შეავსე Timing (ხელით არჩეულ „სხვა“-ზე ცარიელია).")
                    elif not hist_mode and hop_deduct is None:
                        st.error("ერთეულები შეუთავსებელია — ჩანაწერი არ შენახულა, "
                                 "ჩამოჭრა არ მომხდარა.")
                    elif not hist_mode and hop_deduct > hop_stock:
                        st.error(f"მარაგშია მხოლოდ {hop_stock:g} {hop_inv_unit}, "
                                 f"მოთხოვნილია {hop_deduct:g} {hop_inv_unit}.")
                    else:
                        if not hop_existing.empty:
                            new_total = round(_num(hop_existing.iloc[0].get("Qty")) + hop_qty, 4)
                            update_step_qty(bid, "Hop", hop_name, new_total,
                                            timing=hop_timing, brew_day=hop_day)
                        else:
                            append_row("BREW_STEPS", {
                                "Brew_ID": bid, "Stage": "Boil", "Category": "Hop",
                                "Item": hop_name, "Qty": hop_qty, "Unit": "g",
                                "Timing": hop_timing, "AA_%": hop_aa,
                                "Justification": hop_just, "Brew_Day": hop_day,
                            })
                        if not hist_mode:
                            update_cell_by_key("INVENTORY", "Item", hop_name,
                                               "Current_Qty", round(hop_stock - hop_deduct, 4))
                        log_change("სვია დამატება",
                                   f"{hop_name} +{hop_qty:g} g @ {hop_timing}"
                                   + ("" if hist_mode else f" (მარაგი −{hop_deduct:g} {hop_inv_unit})"),
                                   bid=bid, brew_name=choice, day=hop_day)
                        st.rerun()

                if brew_day == 2 and not locked and st.button("📋 დღე 1-ის სვია → დღე 2"):
                    copied, skipped = copy_ingredients_day1_to_day2(bid, "Hop", "g")
                    if copied:
                        st.success(f"დღე 2-ზე დაკოპირდა: {', '.join(copied)} "
                                   "(მარაგიდან ჩამოიჭრა).")
                    if skipped:
                        st.warning(f"მარაგი არ ჰყოფნის, გამოტოვდა: {', '.join(skipped)}.")
                    if copied:
                        st.rerun()

            st.divider()
            st.caption("Pre/Post-Boil და Actual OG — ერთი საერთო მთელ ხარშვაზე (დღეზე არ იყოფა).")
            c1, c2 = st.columns(2)
            pre_boil_vol = c1.number_input("Pre-Boil მოცულობა (L)", disabled=fg_locked)
            pre_boil_p = c2.number_input("Pre-Boil Gravity (°P)", disabled=fg_locked)
            c3, c4 = st.columns(2)
            post_boil_vol = c3.number_input("Post-Boil მოცულობა (L)", disabled=fg_locked)
            actual_og = c4.number_input("Actual OG (°P)", disabled=fg_locked)
            boil_ok = sanity_gate([
                (actual_og > SANITY["og_max_p"],
                 f"Actual OG {actual_og:g}°P — რეალურ დიაპაზონს ({SANITY['og_max_p']:g}°P) სცდება."),
                (pre_boil_p > SANITY["og_max_p"],
                 f"Pre-Boil {pre_boil_p:g}°P — რეალურ დიაპაზონს სცდება."),
                (post_boil_vol and pre_boil_vol and post_boil_vol > pre_boil_vol,
                 f"Post-Boil ({post_boil_vol:g} L) Pre-Boil-ზე ({pre_boil_vol:g} L) მეტია — "
                 f"დუღილისას მოცულობა უნდა შემცირდეს."),
            ], key=f"sanity_boil_{bid}")
            if st.button("💾 boil მონაცემის შენახვა header-ში",
                         disabled=fg_locked or not boil_ok):
                for col, val in [("Pre_Boil_Vol_L", pre_boil_vol), ("Pre_Boil_P", pre_boil_p),
                                  ("Post_Boil_Vol_L", post_boil_vol), ("Actual_OG_P", actual_og)]:
                    update_cell_by_key("BREW_HEADER", "Brew_ID", bid, col, val)
                # OG/post-boil ცვლილებამ ასევე უნდა განაახლოს derived მეტრიკები,
                # მაგრამ მხოლოდ თუ FG დამტკიცებულია (FG_Date შევსებული) — არა
                # უბრალოდ Actual_FG_P-ს არსებობა, რაც შეიძლება დროებითი იყოს
                fg_confirmed = str(row.get("FG_Date", "")).strip() not in ("", "nan", "None")
                if fg_confirmed:
                    recalc_header_metrics(bid, actual_og, row.get("Actual_FG_P"),
                                          post_boil_vol, total_grain_kg(bid),
                                          fg_confirmed=True)
                log_change("boil header",
                           f"PreBoil {pre_boil_vol:g}L/{pre_boil_p:g}°P, "
                           f"PostBoil {post_boil_vol:g}L, OG {actual_og:g}°P",
                           bid=bid, brew_name=choice)
                flash("შენახულია.")
                st.rerun()

            steps_df = get_df("BREW_STEPS")
            if not steps_df.empty:
                boil_view = steps_df[(steps_df["Brew_ID"] == bid) & (steps_df["Stage"] == "Boil")]
                st.dataframe(boil_view, use_container_width=True)

        # ---- FERMENT / GRAVITY TAB ----
        if st.session_state.brew_tab == "🧪 ფერმენტაცია/Gravity":
            locked = lock_gate(row.get("FG_Date"), f"edit_ovr_ferm_{bid}") if fg_done else False

            # --- yeast pitch: which yeast, dry vs slurry, + reuse generation ---
            st.markdown("**🧫 საფუარი (yeast pitch)**")
            yc1, yc2 = st.columns(2)
            yeast_name = yc1.text_input("საფუარი (strain/სახელი)",
                value=str(row.get("Yeast", "") or ""), key=f"yeast_name_{bid}", disabled=locked)
            form_opts = ["ფხვნილი (dry)", "ლექი (slurry)"]
            default_idx = 1 if str(row.get("Yeast_Form", "")).startswith("ლექი") else 0
            yeast_form = yc2.radio("ფორმა", form_opts, index=default_idx,
                horizontal=True, key=f"yeast_form_{bid}", disabled=locked)
            is_slurry = yeast_form.startswith("ლექი")
            if is_slurry:
                # auto-count prior slurry uses of this strain in OTHER brews
                prior = 0
                if yeast_name.strip() and not headers_df.empty and "Yeast" in headers_df.columns:
                    for _, hr in headers_df.iterrows():
                        if (str(hr.get("Brew_ID")) != bid
                                and str(hr.get("Yeast", "")).strip().lower() == yeast_name.strip().lower()
                                and str(hr.get("Yeast_Form", "")).startswith("ლექი")):
                            prior += 1
                saved_gen = int(_num(row.get("Yeast_Generation"), 0))
                # no key on this input: value= always reflects the live suggestion
                yeast_gen = st.number_input(
                    "თაობა (generation) — რამდენჯერ იქნა ეს ლექი გამოყენებული",
                    min_value=1, value=max(1, saved_gen or (prior + 1)), step=1, disabled=locked)
                st.caption(f"ავტო-შეფასება: ამ strain-ის {prior} წინა ლექ-გამოყენება "
                           f"სხვა ხარშვებში → თაობა {prior + 1} (ხელით შეასწორე თუ საჭიროა).")
            else:
                yeast_gen = None
            if st.button("💾 საფუარის შენახვა", disabled=locked):
                form_val = "ლექი" if is_slurry else "ფხვნილი"
                update_cell_by_key("BREW_HEADER", "Brew_ID", bid, "Yeast", yeast_name)
                update_cell_by_key("BREW_HEADER", "Brew_ID", bid, "Yeast_Form", form_val)
                update_cell_by_key("BREW_HEADER", "Brew_ID", bid, "Yeast_Generation",
                                   int(yeast_gen) if is_slurry else "")
                log_change("საფუარი", f"{yeast_name} · {form_val}"
                           + (f" · თაობა {int(yeast_gen)}" if is_slurry else ""),
                           bid=bid, brew_name=choice)
                flash(f"საფუარი შენახულია: {yeast_name} ({form_val}"
                           + (f", თაობა {int(yeast_gen)}" if is_slurry else "") + ").")
                st.rerun()
            st.divider()

            st.markdown("**ყოველდღიური Gravity-ის ჩაწერა**")
            # date + auto Day_# live OUTSIDE the form (form widgets don't
            # recompute until submit, so the day counter must live here)
            gd1, gd2 = st.columns(2)
            g_date = gd1.date_input("თარიღი", value=date.today(), key="g_date", disabled=locked)
            start_raw = str(row.get("Date", "")).strip()
            try:
                start_date = pd.to_datetime(start_raw).date()
            except (ValueError, TypeError):
                start_date = None
            day_number = (g_date - start_date).days if start_date is not None else 0
            gd2.number_input("დღე # (ავტომატური)", value=int(day_number), disabled=True)
            date_error = False
            if start_date is None:
                st.warning("header-ში ხარშვის Date ვერ წავიკითხე — Day_# ვერ დაითვალა.")
            elif day_number < 0:
                date_error = True
                st.error(f"არჩეული თარიღი ({g_date}) ხარშვის დაწყებამდეა "
                         f"({start_date}) — Day_# უარყოფითია. გაასწორე თარიღი.")
            with st.form("gravity_add", clear_on_submit=True):
                c1, c2 = st.columns(2)
                g_raw = c1.number_input("Gravity (°P) — ჰიდრომეტრიდან, დაუკორექტირებელი", disabled=locked)
                g_temp = c2.number_input("ტემპერატურა (°C)", value=10.0, disabled=locked)
                is_final = st.checkbox("ეს არის საბოლოო (FG) გაზომვა", disabled=locked)
                add_g = st.form_submit_button("დამატება", disabled=locked)
                og_now = _num(row.get("Actual_OG_P"))
                if add_g and date_error:
                    st.error("თარიღი ხარშვის დაწყებამდეა — ჩანაწერი არ დაემატა.")
                elif add_g and g_raw > SANITY["og_max_p"]:
                    st.error(f"Gravity {g_raw:g}°P — რეალურ დიაპაზონს "
                             f"({SANITY['og_max_p']:g}°P) სცდება. ჩანაწერი არ დაემატა.")
                elif add_g and og_now and g_raw > og_now:
                    st.error(f"Gravity {g_raw:g}°P OG-ზე ({og_now:g}°P) მეტია — "
                             f"ფერმენტაციისას gravity უნდა ეცემოდეს. ჩანაწერი არ დაემატა.")
                elif add_g:
                    ref_temp = _num(get_setting("Hydrometer_Ref_Temp_C", 20), 20)
                    corrected = correct_gravity_plato(g_raw, g_temp, ref_temp)
                    append_row("BREW_GRAVITY_LOG", {
                        "Brew_ID": bid, "Date": str(g_date), "Day_#": day_number,
                        "Gravity_P_Raw": g_raw, "Temp_C": g_temp, "Instrument": "hydro",
                        "Gravity_P_Corrected": corrected if corrected is not None else "",
                    })
                    if corrected is not None:
                        st.success(f"Corrected: {corrected:.1f}°P (raw {g_raw:.1f}°P @ {g_temp:.0f}°C)")
                    if is_final:
                        fg_value = corrected if corrected is not None else g_raw
                        update_cell_by_key("BREW_HEADER", "Brew_ID", bid, "Actual_FG_P", fg_value)
                        update_cell_by_key("BREW_HEADER", "Brew_ID", bid, "FG_Date", str(g_date))
                        metrics = recalc_header_metrics(
                            bid, row.get("Actual_OG_P"), fg_value,
                            row.get("Post_Boil_Vol_L"), total_grain_kg(bid),
                            fg_confirmed=True)  # checkbox just set FG_Date
                        extra = " | ".join(f"{k} {v}" for k, v in metrics.items())
                        st.success("FG დაფიქსირდა — ხარშვა დახურულია."
                                   + (f" ({extra})" if extra else ""))
                    log_change("FG დაფიქსირდა" if is_final else "Gravity ჩანაწერი",
                               f"day {day_number}: raw {g_raw:g}°P @ {g_temp:g}°C"
                               + (f" → corrected {corrected:g}°P" if corrected is not None else ""),
                               bid=bid, brew_name=choice)
                    st.rerun()

            grav_df = get_df("BREW_GRAVITY_LOG")
            if not grav_df.empty:
                view = grav_df[grav_df["Brew_ID"] == bid].sort_values("Day_#")
                st.dataframe(view, use_container_width=True)
                if len(view) >= 2 and "Gravity_P_Corrected" in view.columns:
                    chart_df = view.set_index("Day_#")[["Gravity_P_Corrected"]] \
                        if view["Gravity_P_Corrected"].notna().any() else view.set_index("Day_#")[["Gravity_P_Raw"]]
                    st.line_chart(chart_df)

            outcome = st.text_area("შედეგის შენიშვნა (გემო/სუნი/სიმღვრივე kegging-ისას)",
                                    value=str(row.get("Outcome_Note", "")), disabled=locked)
            if st.button("💾 შედეგის შენახვა", disabled=locked):
                update_cell_by_key("BREW_HEADER", "Brew_ID", bid, "Outcome_Note", outcome)
                st.success("შენახულია.")

# ============================================================
# PAGE 3 — B2B CLIENTS
# ============================================================
else:
    st.title("🤝 კლიენტები (B2B)")

    clients_df = get_df("CLIENTS")
    products_df = get_df("PRODUCTS")
    ship_df = get_df("SHIPMENTS")
    pay_df = get_df("PAYMENTS")
    moves_df = get_df("ASSET_MOVES")

    # label -> id maps for the selectboxes
    cl_map, pr_map = {}, {}
    if not clients_df.empty and "Client_ID" in clients_df.columns:
        for _, r in clients_df.iterrows():
            if str(r.get("Status", "")).strip() != "არქივი":
                cl_map[str(r.get("Name"))] = r["Client_ID"]
    if not products_df.empty and "Product_ID" in products_df.columns:
        for _, r in products_df.iterrows():
            pr_map[f"{r.get('Name')} ({r.get('Category')})"] = r["Product_ID"]

    CL_TABS = ["📋 მიმოხილვა", "👤 კლიენტის ბარათი", "🚚 გატანა",
               "💰 გადახდა", "📦 ინვენტარი", "⚙️ კლიენტები/პროდუქტები"]
    if "cl_tab" not in st.session_state:
        st.session_state.cl_tab = "📋 მიმოხილვა"
    st.session_state.cl_tab = st.radio(
        "ტაბი", CL_TABS, horizontal=True, label_visibility="collapsed",
        key="cl_tab_radio", index=CL_TABS.index(st.session_state.cl_tab))

    # ---------------- OVERVIEW ----------------
    if st.session_state.cl_tab == "📋 მიმოხილვა":
        if not cl_map:
            st.info("კლიენტი ჯერ არ არის დამატებული — იხ. ტაბი „⚙️ კლიენტები/პროდუქტები“.")
        else:
            rows = []
            for name, cid in cl_map.items():
                bal = client_balance(cid, ship_df, pay_df)
                assets = client_assets(cid, moves_df)
                last = ""
                if not ship_df.empty and "Client_ID" in ship_df.columns:
                    s = ship_df[ship_df["Client_ID"] == cid]
                    if not s.empty:
                        last = str(s.iloc[-1].get("Date", ""))
                other = ", ".join(f"{k} {v:g}" for k, v in assets.items() if k != "კეგი")
                rows.append({
                    "კლიენტი": name,
                    "დავალიანება (₾)": bal,
                    "კეგი": assets.get("კეგი", 0),
                    "სხვა ინვენტარი": other or "—",
                    "ბოლო გატანა": last or "—",
                })
            ov = pd.DataFrame(rows).sort_values("დავალიანება (₾)", ascending=False)
            c1, c2, c3, c4 = st.columns(4, gap="large")
            c1.metric("სულ დავალიანება (₾)", f"{ov['დავალიანება (₾)'].sum():g}")
            c2.metric("კლიენტები", len(ov))
            c3.metric("კეგი კლიენტებთან", f"{kegs_at_clients(moves_df):g}")
            c4.metric("ქარხანაში თავისუფალი", f"{kegs_free():g}")
            st.caption(f"სულ კეგი: {kegs_total():g} (შესაცვლელად — SETTINGS → Kegs_Total)")
            st.dataframe(ov, use_container_width=True, hide_index=True)

    # ---------------- CLIENT CARD ----------------
    elif st.session_state.cl_tab == "👤 კლიენტის ბარათი":
        if not cl_map:
            st.info("ჯერ დაამატე კლიენტი.")
        else:
            pick = st.selectbox("კლიენტი", list(cl_map.keys()), key="card_client")
            cid = cl_map[pick]
            crow = clients_df[clients_df["Client_ID"] == cid].iloc[0]
            bal = client_balance(cid, ship_df, pay_df)
            assets = client_assets(cid, moves_df)

            with st.container(border=True, key="clcard-head"):
                c1, c2, c3 = st.columns(3)
                c1.markdown(f"**ტიპი**  \n{crow.get('Type') or '—'}")
                c2.markdown(f"**საკონტაქტო**  \n{crow.get('Contact_Person') or '—'}")
                c3.markdown(f"**ტელეფონი**  \n{crow.get('Phone') or '—'}")
                st.metric("დავალიანება (₾)", f"{bal:g}")
                if assets:
                    st.markdown("**მასთან არსებული ინვენტარი:** "
                                + " · ".join(f"{k} **{v:g}**" for k, v in assets.items()))
                else:
                    st.caption("ინვენტარი მასთან არ ირიცხება.")

            with st.container(border=True, key="clcard-ship"):
                st.markdown("#### 🚚 გატანების ისტორია")
                s = ship_df[ship_df["Client_ID"] == cid] if not ship_df.empty else pd.DataFrame()
                if s.empty:
                    st.caption("გატანა ჯერ არ არის.")
                else:
                    s = s.copy()
                    s["პროდუქტი"] = s["Product_ID"].map(
                        lambda p: next((k for k, v in pr_map.items() if v == p), p))
                    cols = ["Date", "პროდუქტი", "Volume_L", "Price_per_L", "Total_GEL",
                            "Paid_Now_GEL", "Kegs_Out", "Notes"]
                    st.dataframe(s[[c for c in cols if c in s.columns]].iloc[::-1],
                                 use_container_width=True, hide_index=True)

            with st.container(border=True, key="clcard-pay"):
                st.markdown("#### 💰 გადახდები")
                p = pay_df[pay_df["Client_ID"] == cid] if not pay_df.empty else pd.DataFrame()
                if p.empty:
                    st.caption("ცალკე გადახდა ჯერ არ არის.")
                else:
                    cols = ["Date", "Amount_GEL", "Method", "Notes"]
                    st.dataframe(p[[c for c in cols if c in p.columns]].iloc[::-1],
                                 use_container_width=True, hide_index=True)

            with st.container(border=True, key="clcard-assets"):
                st.markdown("#### 📦 ინვენტარის მოძრაობა")
                m = moves_df[moves_df["Client_ID"] == cid] if not moves_df.empty else pd.DataFrame()
                if m.empty:
                    st.caption("მოძრაობა ჯერ არ არის.")
                else:
                    cols = ["Date", "Asset_Type", "Detail", "Direction", "Qty", "Notes"]
                    st.dataframe(m[[c for c in cols if c in m.columns]].iloc[::-1],
                                 use_container_width=True, hide_index=True)

    # ---------------- SHIPMENT ----------------
    elif st.session_state.cl_tab == "🚚 გატანა":
        if not cl_map or not pr_map:
            st.info("ჯერ დაამატე კლიენტი და პროდუქტი („⚙️“ ტაბი).")
        else:
            c1, c2 = st.columns(2)
            cname = c1.selectbox("კლიენტი", list(cl_map.keys()), key="sh_client")
            pname = c2.selectbox("პროდუქტი", list(pr_map.keys()), key="sh_product")
            cid, pid = cl_map[cname], pr_map[pname]

            ksize = keg_size_l()
            held_now = client_assets(cid, moves_df).get("კეგი", 0)
            free_now = kegs_free()
            k1, k2 = st.columns(2)
            k1.metric("ქარხანაში თავისუფალი კეგი", f"{free_now:g}")
            k2.metric(f"{cname}-თან ახლა", f"{held_now:g}")

            cur_price = client_price(cid, pid, get_df("CLIENT_PRICES"), products_df)
            # in a form, so typing a value and hitting save works on the FIRST
            # click (outside a form the first click only commits the field)
            with st.form("b2b_shipment", clear_on_submit=True):
                c3, c4, c5 = st.columns(3)
                s_date = c3.date_input("თარიღი", value=date.today())
                kegs = c4.number_input(f"წაიღო სავსე კეგი (×{ksize:g}ლ)",
                                       min_value=0, step=1)
                price = c5.number_input("ფასი ₾/ლ", min_value=0.0, step=0.1,
                                        value=float(cur_price))
                c6, c7 = st.columns(2)
                kegs_back = c6.number_input("დააბრუნა ცარიელი კეგი", min_value=0, step=1)
                # never assume payment: the debt is created by default and any
                # payment is entered explicitly
                paid_now = c7.number_input("ახლა გადაიხადა (₾)", min_value=0.0,
                                           step=1.0, value=0.0)
                note = st.text_input("შენიშვნა")
                save_price = st.checkbox(
                    f"შეცვლილი ფასი დაიმახსოვრე ამ კლიენტისთვის (ახლა {cur_price:g} ₾/ლ)",
                    value=True)
                confirm_odd = st.checkbox("დიახ, უჩვეულო მონაცემიც სწორია")
                submitted = st.form_submit_button("💾 გატანის ჩაწერა")

            if submitted:
                vol = kegs * ksize
                total = round(vol * price, 2)
                problems = [m for bad, m in [
                    (price > 100, f"ფასი {price:g} ₾/ლ — ჩვეულებრივზე ბევრად მეტია."),
                    (paid_now > total, f"გადახდილი ({paid_now:g} ₾) ჯამზე ({total:g} ₾) მეტია."),
                    (kegs > free_now,
                     f"წასაღებია {kegs} კეგი, ქარხანაში კი {free_now:g} თავისუფალია."),
                    (kegs_back > held_now,
                     f"აბრუნებს {kegs_back} კეგს, მაგრამ მასთან {held_now:g} ირიცხება."),
                ] if bad]
            else:
                vol = total = 0
                problems = []

            if submitted and kegs <= 0 and kegs_back <= 0:
                st.error("შეავსე კეგების რაოდენობა.")
            elif submitted and problems and not confirm_odd:
                for m in problems:
                    st.warning(f"⚠️ {m}")
                st.error("დაადასტურე ქვედა გრაფა („უჩვეულო მონაცემიც სწორია“) და ხელახლა შეინახე.")
            elif submitted:
                sid = new_id("SHIP")
                append_row("SHIPMENTS", {
                    "Shipment_ID": sid, "Date": str(s_date), "Client_ID": cid,
                    "Product_ID": pid, "Volume_L": vol, "Price_per_L": price,
                    "Total_GEL": total, "Paid_Now_GEL": paid_now, "Kegs_Out": kegs,
                    "Kegs_Returned": kegs_back, "Notes": note,
                    "Operator": st.session_state.get("operator", ""),
                })
                op = st.session_state.get("operator", "")
                if kegs_back:  # empties come back first
                    append_row("ASSET_MOVES", {
                        "Move_ID": new_id("MOVE"), "Date": str(s_date), "Client_ID": cid,
                        "Asset_Type": "კეგი", "Detail": "ცარიელი", "Direction": "დაბრუნება",
                        "Qty": kegs_back, "Notes": f"გატანასთან ერთად ({sid})", "Operator": op,
                    })
                if kegs:  # then the full ones go out
                    append_row("ASSET_MOVES", {
                        "Move_ID": new_id("MOVE"), "Date": str(s_date), "Client_ID": cid,
                        "Asset_Type": "კეგი", "Detail": "სავსე", "Direction": "გატანა",
                        "Qty": kegs, "Notes": f"გატანასთან ერთად ({sid})", "Operator": op,
                    })
                if save_price and abs(price - cur_price) > 1e-9:
                    set_client_price(cid, pid, price)
                log_change("B2B გატანა",
                           f"{cname}: {pname} {kegs} კეგი ({vol:g}ლ) × {price:g}₾ = {total:g}₾, "
                           f"გადახდილი {paid_now:g}₾, დააბრუნა {kegs_back}")
                rest = total - paid_now
                flash(f"ჩაწერილია: {total:g} ₾"
                           + (f", ნაშთი {rest:g} ₾" if rest > 0 else ", სრულად გადახდილი"))
                st.rerun()

    # ---------------- PAYMENT ----------------
    elif st.session_state.cl_tab == "💰 გადახდა":
        if not cl_map:
            st.info("ჯერ დაამატე კლიენტი.")
        else:
            cname = st.selectbox("კლიენტი", list(cl_map.keys()), key="pay_client")
            cid = cl_map[cname]
            bal = client_balance(cid, ship_df, pay_df)
            st.metric("მიმდინარე დავალიანება (₾)", f"{bal:g}")
            with st.form("b2b_payment", clear_on_submit=True):
                c1, c2, c3 = st.columns(3)
                p_date = c1.date_input("თარიღი", value=date.today())
                amount = c2.number_input("თანხა (₾)", min_value=0.0, step=10.0,
                                         value=float(max(bal, 0)))
                method = c3.selectbox("მეთოდი", ["ნაღდი", "გადარიცხვა", "სხვა"])
                note = st.text_input("შენიშვნა")
                confirm_over = st.checkbox("დიახ, დავალიანებაზე მეტია (წინსწრებით)")
                submitted = st.form_submit_button("💾 გადახდის ჩაწერა")

            if submitted and amount <= 0:
                st.error("შეავსე თანხა.")
            elif submitted and amount > bal + 0.01 and bal > 0 and not confirm_over:
                st.warning(f"⚠️ გადახდა ({amount:g} ₾) დავალიანებაზე ({bal:g} ₾) მეტია.")
                st.error("დაადასტურე ქვედა გრაფა და ხელახლა შეინახე.")
            elif submitted:
                append_row("PAYMENTS", {
                    "Payment_ID": new_id("PAY"), "Date": str(p_date), "Client_ID": cid,
                    "Amount_GEL": amount, "Method": method, "Shipment_ID": "",
                    "Notes": note, "Operator": st.session_state.get("operator", ""),
                })
                log_change("B2B გადახდა", f"{cname}: {amount:g}₾ ({method})")
                flash(f"ჩაწერილია. ახალი ბალანსი: {bal - amount:g} ₾")
                st.rerun()

    # ---------------- ASSETS ----------------
    elif st.session_state.cl_tab == "📦 ინვენტარი":
        if not cl_map:
            st.info("ჯერ დაამატე კლიენტი.")
        else:
            cname = st.selectbox("კლიენტი", list(cl_map.keys()), key="as_client")
            cid = cl_map[cname]
            held = client_assets(cid, moves_df)
            k1, k2 = st.columns(2)
            k1.metric("ქარხანაში თავისუფალი კეგი", f"{kegs_free():g}")
            k2.metric("კეგი ამ კლიენტთან", f"{held.get('კეგი', 0):g}")
            if held:
                st.markdown("**ამჟამად მასთან:** "
                            + " · ".join(f"{k} **{v:g}**" for k, v in held.items()))
            else:
                st.caption("ამჟამად ინვენტარი მასთან არ ირიცხება.")

            with st.form("b2b_asset", clear_on_submit=True):
                c1, c2 = st.columns(2)
                atype_pick = c1.selectbox("ტიპი", ASSET_TYPES + ["✏️ სხვა (ხელით)"])
                atype_custom = c1.text_input("ახალი ტიპი (თუ „სხვა“ აირჩიე)")
                direction = c2.radio("მიმართულება", ["გატანა", "დაბრუნება"], horizontal=True)
                c3, c4, c5 = st.columns(3)
                a_date = c3.date_input("თარიღი", value=date.today())
                qty = c4.number_input("რაოდენობა", min_value=1, step=1)
                detail = c5.text_input("დეტალი (მაგ. 50ლ, სერიული)")
                note = st.text_input("შენიშვნა")
                confirm_over = st.checkbox("დიახ, მაინც ჩაწერე")
                submitted = st.form_submit_button("💾 ჩაწერა")

            atype = atype_custom.strip() if atype_pick.startswith("✏️") else atype_pick
            have = held.get(atype, 0)
            if submitted and not atype:
                st.error("შეავსე ტიპის დასახელება.")
            elif (submitted and direction == "დაბრუნება" and qty > have
                  and not confirm_over):
                st.warning(f"⚠️ დაბრუნება {qty} ცალი, მაგრამ მასთან {have:g} ირიცხება.")
                st.error("დაადასტურე ქვედა გრაფა და ხელახლა შეინახე.")
            elif submitted:
                append_row("ASSET_MOVES", {
                    "Move_ID": new_id("MOVE"), "Date": str(a_date), "Client_ID": cid,
                    "Asset_Type": atype, "Detail": detail, "Direction": direction,
                    "Qty": qty, "Notes": note,
                    "Operator": st.session_state.get("operator", ""),
                })
                log_change("B2B ინვენტარი", f"{cname}: {atype} {direction} {qty} ცალი")
                flash(f"{atype} — {direction} {qty} ცალი ჩაწერილია.")
                st.rerun()

    # ---------------- SETUP ----------------
    else:
        st.markdown("### 👥 კლიენტები")
        if not clients_df.empty:
            st.dataframe(clients_df, use_container_width=True, hide_index=True)
        with st.form("add_client", clear_on_submit=True):
            st.markdown("**ახალი კლიენტი**")
            c1, c2, c3 = st.columns(3)
            cname_new = c1.text_input("დასახელება")
            ctype = c2.selectbox("ტიპი", ["ბარი", "რესტორანი", "მაღაზია", "სხვა"])
            cperson = c3.text_input("საკონტაქტო პირი")
            c4, c5, c6 = st.columns(3)
            cphone = c4.text_input("ტელეფონი")
            caddr = c5.text_input("მისამართი")
            cterms = c6.number_input("გადახდის ვადა (დღე)", min_value=0, step=1, value=0)
            cnote = st.text_input("შენიშვნა", key="cl_note")
            if st.form_submit_button("დამატება") and cname_new:
                append_row("CLIENTS", {
                    "Client_ID": new_id("CL"), "Name": cname_new, "Type": ctype,
                    "Contact_Person": cperson, "Phone": cphone, "Address": caddr,
                    "Status": "აქტიური", "Payment_Terms_Days": cterms,
                    "Created_Date": str(date.today()), "Notes": cnote,
                })
                log_change("B2B კლიენტი დაემატა", cname_new)
                flash(f"{cname_new} დაემატა.")
                st.rerun()

        st.divider()
        st.markdown("### 🍾 პროდუქტები")
        if not products_df.empty:
            st.dataframe(products_df, use_container_width=True, hide_index=True)
        with st.form("add_product", clear_on_submit=True):
            st.markdown("**ახალი პროდუქტი**")
            c1, c2, c3 = st.columns(3)
            pname_new = c1.text_input("დასახელება")
            pcat = c2.selectbox("კატეგორია", PRODUCT_CATEGORIES)
            pprice = c3.number_input("ბაზისური ფასი ₾/ლ", min_value=0.0, step=0.1)
            pnote = st.text_input("შენიშვნა", key="pr_note")
            if st.form_submit_button("დამატება") and pname_new:
                append_row("PRODUCTS", {
                    "Product_ID": new_id("PR"), "Name": pname_new, "Category": pcat,
                    "Unit": "ლ", "Default_Price_GEL": pprice, "Active": "1",
                    "Notes": pnote,
                })
                log_change("B2B პროდუქტი დაემატა", f"{pname_new} ({pcat}) {pprice:g}₾/ლ")
                flash(f"{pname_new} დაემატა.")
                st.rerun()

        if cl_map and pr_map:
            st.divider()
            st.markdown("### 💵 კლიენტის ფასები")
            c1, c2, c3 = st.columns(3)
            pc = c1.selectbox("კლიენტი", list(cl_map.keys()), key="pp_client")
            pp = c2.selectbox("პროდუქტი", list(pr_map.keys()), key="pp_product")
            cur = client_price(cl_map[pc], pr_map[pp], get_df("CLIENT_PRICES"), products_df)
            newp = c3.number_input("ფასი ₾/ლ", min_value=0.0, step=0.1,
                                   value=float(cur), key="pp_price")
            if st.button("💾 ფასის შენახვა"):
                set_client_price(cl_map[pc], pr_map[pp], newp)
                log_change("B2B ფასი", f"{pc} / {pp}: {cur:g} → {newp:g} ₾/ლ")
                flash("ფასი შენახულია.")
                st.rerun()
