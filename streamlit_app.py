import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime, date
import uuid

st.set_page_config(page_title="BUGHASHVILI Brew Journal", layout="wide")

# ============================================================
# GLOBAL DESIGN SYSTEM — one CSS block, loaded once for the whole app.
# Colors/typography come from .streamlit/config.toml; this only covers
# what config can't express. Selectors are stable: data-testid (Streamlit
# public test hooks) and .st-key-* (from widget key=), NOT version-specific
# emotion hashes.
# ============================================================
st.markdown("""
<style>
/* card polish: subtle warm fill on overview cards (keyed = version-stable) */
[class*="st-key-ovcard-"] {
    background: rgba(224, 162, 60, 0.04);
}
/* emphasized numbers (OG/FG/Eff/ABV) — big, beer-amber */
[data-testid="stMetricValue"] {
    font-size: 1.9rem;
    font-weight: 700;
    color: #E9A93C;
}
[data-testid="stMetricLabel"] p {
    font-size: 0.78rem;
    letter-spacing: 0.04em;
    opacity: 0.72;
}
/* brew page: horizontal radio styled as tabs (moved here from inline
   so it loads once, not on every brew-page render) */
.st-key-brew_tab_radio div[role="radiogroup"] { gap: 4px; }
.st-key-brew_tab_radio div[role="radiogroup"] label {
    border: 1px solid rgba(224, 162, 60, 0.30);
    border-bottom: none;
    border-radius: 10px 10px 0 0;
    padding: 6px 16px;
    background: rgba(255, 255, 255, 0.02);
}
.st-key-brew_tab_radio div[role="radiogroup"] label:has(input:checked) {
    background: rgba(224, 162, 60, 0.16);
    border-color: rgba(224, 162, 60, 0.6);
}
.st-key-brew_tab_radio div[role="radiogroup"] label > div:first-child { display: none; }
/* sidebar nav: highlight the selected page */
.st-key-nav_radio div[role="radiogroup"] label {
    padding: 6px 10px;
    border-radius: 8px;
    margin-bottom: 2px;
}
.st-key-nav_radio div[role="radiogroup"] label:has(input:checked) {
    background: rgba(224, 162, 60, 0.15);
    box-shadow: inset 3px 0 0 #E0A23C;
}

/* ---------- mobile / touch ---------- */
/* touch-friendly targets everywhere (Apple HIG minimum 44px) */
.stButton button, .stFormSubmitButton button, .stDownloadButton button {
    min-height: 44px;
}
[data-testid="stNumberInput"] input,
[data-testid="stTextInput"] input,
[data-testid="stDateInput"] input,
[data-baseweb="select"] > div {
    min-height: 44px;
}
[data-testid="stNumberInput"] button { min-width: 40px; }  /* +/- steppers */
/* widget labels: wrap instead of truncate on narrow screens */
[data-testid="stWidgetLabel"] p {
    white-space: normal;
    overflow-wrap: anywhere;
}
/* sidebar toggle (hamburger when collapsed): bigger, amber, obvious.
   stExpandSidebarButton = 1.59 testid; stSidebarCollapsedControl = older */
button[data-testid="stExpandSidebarButton"],
[data-testid="stSidebarCollapsedControl"] {
    background: rgba(224, 162, 60, 0.15) !important;
    border: 1px solid rgba(224, 162, 60, 0.5) !important;
    border-radius: 10px;
    min-height: 44px;
    min-width: 44px;
}
button[data-testid="stExpandSidebarButton"] *,
[data-testid="stSidebarCollapsedControl"] * { color: #E0A23C !important; }

@media (max-width: 640px) {
    /* primary action buttons stretch to full width on phones.
       the wrapper divs are fit-content, so widen the whole chain */
    [data-testid="stElementContainer"]:has([data-testid="stButton"]),
    [data-testid="stElementContainer"]:has([data-testid="stFormSubmitButton"]) { width: 100%; }
    .stButton, .stFormSubmitButton { width: 100%; }
    .stButton button, .stFormSubmitButton button { width: 100%; }
    /* brew tab bar: single row, horizontally swipeable */
    .st-key-brew_tab_radio div[role="radiogroup"] {
        flex-wrap: nowrap;
        overflow-x: auto;
        -webkit-overflow-scrolling: touch;
        scrollbar-width: none;
        padding-bottom: 2px;
    }
    .st-key-brew_tab_radio div[role="radiogroup"]::-webkit-scrollbar { display: none; }
    .st-key-brew_tab_radio div[role="radiogroup"] label {
        flex: 0 0 auto;
        white-space: nowrap;
    }
    /* slightly tighter page padding on phones */
    [data-testid="stMainBlockContainer"] { padding-left: 1rem; padding-right: 1rem; }
}
</style>
""", unsafe_allow_html=True)

# ============================================================
# CONNECTION — reads credentials + sheet id from Streamlit secrets
# Local: .streamlit/secrets.toml
# Cloud: App settings -> Secrets (same format, pasted in)
# ============================================================
SCOPES = ["https://www.googleapis.com/auth/spreadsheets",
          "https://www.googleapis.com/auth/drive"]

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

def ws(name):
    """Get a worksheet, creating it with headers if missing (except core sheets)."""
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
def get_df(name):
    w = ws(name)
    data = w.get_all_records()
    if not data:
        # empty sheet: keep the real columns so downstream df["col"]
        # filters don't KeyError
        headers = w.row_values(1)
        return pd.DataFrame(columns=headers)
    return pd.DataFrame(data)

def append_row(name, row_dict):
    """Append a row, aligning values to the sheet's existing header order."""
    w = ws(name)
    headers = w.row_values(1)
    row = [row_dict.get(h, "") for h in headers]
    w.append_row(row, value_input_option="USER_ENTERED")
    get_df.clear()  # invalidate cache so the new row shows immediately

def update_cell_by_key(name, key_col, key_val, target_col, value):
    w = ws(name)
    headers = w.row_values(1)
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
}

def new_id(prefix):
    return f"{prefix}-{datetime.now().strftime('%Y%m%d-%H%M%S')}-{uuid.uuid4().hex[:4]}"

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
    """Convert qty from the form's unit to the INVENTORY row's unit via a
    common base (grams/milliliters). Returns None when the units are
    incompatible (e.g. 'pack' vs 'kg') — caller must block and warn instead
    of computing blindly."""
    fu = str(from_unit).strip().lower()
    iu = str(inventory_unit).strip().lower()
    if fu not in UNIT_TO_GRAMS_OR_ML or iu not in UNIT_TO_GRAMS_OR_ML:
        return None
    if _UNIT_GROUP[fu] != _UNIT_GROUP[iu]:
        return None
    return qty * UNIT_TO_GRAMS_OR_ML[fu] / UNIT_TO_GRAMS_OR_ML[iu]

def find_inventory_item(inv_df, patterns):
    """Find an INVENTORY row whose Item name contains any of the (lowercase)
    patterns. Returns the row (Series) or None. Name-tolerant on purpose —
    e.g. 'GYpsum' matches 'gypsum'."""
    if inv_df.empty or "Item" not in inv_df.columns:
        return None
    for _, r in inv_df.iterrows():
        name = str(r.get("Item", "")).lower()
        if any(p in name for p in patterns):
            return r
    return None

def delete_rows_where(sheet_name, col_name, value):
    """Delete every row in a worksheet whose col_name equals value.
    Collects 1-based row indices first, then deletes bottom-up so indices
    don't shift. Returns the number of deleted rows."""
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
    """Single source of truth for 'finished': FG_Date is set (i.e. the
    'ეს არის FG' checkbox was used). Used by the edit-lock gate."""
    match = headers_df[headers_df["Brew_ID"] == brew_id]
    if match.empty:
        return False
    return str(match.iloc[0].get("FG_Date", "")).strip() not in ("", "nan", "None")

def lock_gate(fg_date, key):
    """Render the finished-brew warning + override checkbox for a tab.
    Returns True when the tab should stay locked (inputs disabled)."""
    st.warning(f"ეს ხარშვა დასრულებულია (FG დაფიქსირდა {fg_date}-ზე). "
               f"ცვლილება არ არის რეკომენდებული.")
    return not st.checkbox("მაინც მინდა რედაქტირება", key=key)

def update_step_qty(bid, category, item, new_qty, timing=None):
    """Update the Qty of the BREW_STEPS row matching Brew_ID + Category + Item
    (a multi-column match — update_cell_by_key can't do that, it keys on one
    column). When timing is given (hops), the Timing column must match too:
    the same hop at a different timing is a separate addition, not a
    duplicate. Returns True if a row was found and updated."""
    w = ws("BREW_STEPS")
    values = w.get_all_values()
    if not values:
        return False
    headers = values[0]
    try:
        bi, ci, ii, qi = (headers.index("Brew_ID"), headers.index("Category"),
                          headers.index("Item"), headers.index("Qty"))
        ti = headers.index("Timing") if timing is not None else None
    except ValueError:
        return False
    for ridx, rowv in enumerate(values[1:], start=2):
        if (len(rowv) > max(bi, ci, ii) and rowv[bi] == bid
                and rowv[ci] == category and rowv[ii] == item
                and (ti is None or (len(rowv) > ti and rowv[ti] == str(timing)))):
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
    """Water density-correction polynomial used by the sheet's hydrometer
    temperature correction (t in °C)."""
    return (1.00130346 - 0.000134722124 * t + 0.00000204052596 * t ** 2
            - 0.00000000232820948 * t ** 3)

def correct_gravity_plato(raw_p, temp_c, ref_temp_c):
    """Port of the BREW_GRAVITY_LOG Gravity_P_Corrected sheet formula.
    Converts a raw °P hydrometer reading at temp_c to °P corrected to the
    hydrometer's reference temperature. Returns rounded float, or None if
    inputs are missing."""
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
    """Compute BREW_HEADER derived metrics in Python (formerly sheet formulas
    that only existed on rows 2-22). Returns only the metrics whose inputs
    are present."""
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
    """Recompute and write Total_Grain_kg + the three derived metrics into
    BREW_HEADER for one brew. og/fg/post_boil_vol/grain_kg are coerced; missing
    inputs simply skip the metrics they feed.

    Eff/ADF/ABV are FG-dependent and only meaningful once FG is *confirmed*
    (the "ეს არის FG" checkbox, i.e. FG_Date is set). Pass fg_confirmed=True
    only in that case; otherwise those three are skipped so a provisional
    interim reading never lands in the header. Total_Grain_kg is written
    regardless — it is valid independently of FG."""
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
page = st.sidebar.radio("გვერდი", ["📦 Inventory", "🍺 ხარშვა"], key="nav_radio")

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
            headers = w.row_values(1)
            full = df.copy()
            full.update(edited)
            values = [headers] + full[headers].astype(str).values.tolist()
            w.clear()
            w.update(values)
            get_df.clear()
            st.success("შენახულია.")
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
            st.success(f"{item} დაემატა.")
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
                st.success(f"წაშლილია ({n} row).")
                st.rerun()

# ============================================================
# PAGE 2 — BREW
# ============================================================
else:
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
        with st.form("new_brew"):
            c1, c2, c3 = st.columns(3)
            b_date = c1.date_input("თარიღი", value=date.today())
            style = c2.text_input("ლუდის სტილი", placeholder="მაგ. Märzen/Oktoberfest")
            ferm = c3.selectbox("ფერმენტორი", ["CCT1", "CCT2"])
            c4, c5 = st.columns(2)
            target_vol = c4.number_input("სამიზნე მოცულობა (L)", min_value=0.0, value=850.0)
            water_us = c5.number_input("წყლის Water_uS", min_value=0.0, value=70.0)
            c6, c7 = st.columns(2)
            target_og = c6.number_input("Target OG (°P)", min_value=0.0, step=0.1)
            target_fg = c7.number_input("Target FG (°P)", min_value=0.0, step=0.1)
            start = st.form_submit_button("ხარშვის დაწყება")
            if start and style:
                bid = new_id("BREW")
                # +2 offset: brews 1 & 2 were lost, so numbering starts higher
                # to stay consistent with the renumbered existing brews
                display_name = f"ხარშვა {len(headers_df) + 1 + 2} — {style}"
                append_row("BREW_HEADER", {
                    "Brew_ID": bid, "Date": str(b_date), "Beer_Style": style,
                    "Fermenter": ferm, "Target_Vol_L": target_vol, "Water_uS": water_us,
                    "Target_OG_P": target_og, "Target_FG_P": target_fg,
                    "Display_Name": display_name,
                })
                st.success(f"ხარშვა შეიქმნა: {display_name} ({bid}). "
                           f"აირჩიე ის ზემოთა სიიდან რომ დეტალები შეავსო.")
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
                c1, c2, c3 = st.columns(3, gap="large")
                c1.metric("Total Gypsum (g)", _show(row.get("Total_Gypsum_g")))
                c2.metric("Total CaCl₂ (g)", _show(row.get("Total_CaCl2_g")))
                c3.metric("Total Lactic (ml)", _show(row.get("Total_Lactic_ml")))
                if not water_b.empty:
                    wcols = [c for c in ["Water_Stream", "Volume_L", "Gypsum_g", "CaCl2_g",
                                         "Lactic_ml", "Target_pH"] if c in water_b.columns]
                    st.dataframe(water_b[wcols], use_container_width=True, hide_index=True)
                else:
                    st.caption("წყლის მონაცემი ჯერ არ არის შენახული.")

            # --- grain bill card ---
            with st.container(border=True, key="ovcard-grain"):
                st.markdown("#### 🌾 ალაო (grain bill)")
                malt_b = (steps_b[steps_b["Category"] == "Malt"]
                          if not steps_b.empty and "Category" in steps_b.columns else pd.DataFrame())
                if not malt_b.empty:
                    mcols = [c for c in ["Item", "Qty", "Unit", "Justification"] if c in malt_b.columns]
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
                    mscols = [c for c in ["Item", "Qty", "Unit", "Timing", "Justification"]
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
                    hcols = [c for c in ["Item", "Qty", "Unit", "Timing", "AA_%", "Justification"]
                             if c in hop_b.columns]
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
                        st.success("წაშლილია: " + ", ".join(
                            f"{s} −{n}" for s, n in deleted.items()))
                        st.rerun()

        # ---- WATER TAB ----
        if st.session_state.brew_tab == "💧 წყალი":
            locked = lock_gate(row.get("FG_Date"), f"edit_ovr_water_{bid}") if fg_done else False
            profiles_df = get_df("WATER_PROFILES")
            profile_names = sorted(profiles_df["Profile_Name"].unique()) if not profiles_df.empty else []
            chosen_profile = st.selectbox("წყლის პროფილი (თუ შენახული გაქვს)",
                                           ["— ხელით შევსება —"] + profile_names, disabled=locked)

            defaults = {"Mash": {"vol": None, "gyp": 0, "cacl2": 0, "lac": 0, "ph": 5.3},
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

            st.markdown("**Mash წყალი** _(მოცულობა ყოველ ხარშვაზე იცვლება — შეავსე ხელით)_")
            m1, m2, m3, m4, m5 = st.columns(5)
            m_vol = m1.number_input("მოცულობა (L)", min_value=0.0, key="m_vol", disabled=locked)
            m_gyp = m2.number_input("Gypsum (g)", value=float(defaults["Mash"]["gyp"] or 0), key="m_gyp", disabled=locked)
            m_cacl = m3.number_input("CaCl2 (g)", value=float(defaults["Mash"]["cacl2"] or 0), key="m_cacl", disabled=locked)
            m_lac = m4.number_input("Lactic (ml)", value=float(defaults["Mash"]["lac"] or 0), key="m_lac", disabled=locked)
            m_ph = m5.number_input("Target pH", value=float(defaults["Mash"]["ph"] or 5.3), key="m_ph", disabled=locked)

            st.markdown("**Sparge წყალი**")
            s1, s2, s3, s4, s5 = st.columns(5)
            s_vol = s1.number_input("მოცულობა (L)", value=float(defaults["Sparge"]["vol"] or 1000), key="s_vol", disabled=locked)
            s_gyp = s2.number_input("Gypsum (g)", value=float(defaults["Sparge"]["gyp"] or 0), key="s_gyp", disabled=locked)
            s_cacl = s3.number_input("CaCl2 (g)", value=float(defaults["Sparge"]["cacl2"] or 0), key="s_cacl", disabled=locked)
            s_lac = s4.number_input("Lactic (ml)", value=float(defaults["Sparge"]["lac"] or 0), key="s_lac", disabled=locked)
            s_ph = s5.number_input("Target pH", value=float(defaults["Sparge"]["ph"] or 5.7), key="s_ph", disabled=locked)

            if st.button("💾 წყლის მონაცემის შენახვა", disabled=locked):
                # salt auto-deduction: total need across Mash+Sparge, in grams.
                # Only NEW saves deduct — existing WATER_TREATMENT rows are
                # never touched retroactively.
                need = {"Gypsum": m_gyp + s_gyp, "CaCl2": m_cacl + s_cacl}
                blocked = False
                deductions = {}  # salt -> (item_name, stock, deduct_in_inv_unit, inv_unit)
                if not hist_mode:  # historical entries never validate/deduct stock
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
                    append_row("WATER_TREATMENT", {"Brew_ID": bid, "Water_Stream": "Mash",
                        "Volume_L": m_vol, "Gypsum_g": m_gyp, "CaCl2_g": m_cacl,
                        "Lactic_ml": m_lac, "Target_pH": m_ph})
                    append_row("WATER_TREATMENT", {"Brew_ID": bid, "Water_Stream": "Sparge",
                        "Volume_L": s_vol, "Gypsum_g": s_gyp, "CaCl2_g": s_cacl,
                        "Lactic_ml": s_lac, "Target_pH": s_ph})
                    for salt, (item, stock, deduct, inv_unit) in deductions.items():
                        update_cell_by_key("INVENTORY", "Item", item,
                                           "Current_Qty", round(stock - deduct, 4))
                    msg = "წყლის მონაცემი შენახულია."
                    if hist_mode:
                        msg += " 📜 ისტორიული რეჟიმი — მარაგი უცვლელია."
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
                st.success(f"პროფილი '{new_profile_name}' შენახულია.")
                st.rerun()

            existing_water = get_df("WATER_TREATMENT")
            if not existing_water.empty:
                st.dataframe(existing_water[existing_water["Brew_ID"] == bid],
                             use_container_width=True)

        # ---- MASH TAB ----
        if st.session_state.brew_tab == "🌾 მეშინგი":
            locked = lock_gate(row.get("FG_Date"), f"edit_ovr_mash_{bid}") if fg_done else False
            st.caption("[Certain] mash schedule ხარშვიდან ხარშვამდე იცვლება — ამიტომ ყოველთვის ხელით.")
            st.markdown("**საფეხურები** _(შეავსე ის row-ები, რომლებიც გჭირდება — ცარიელი row-ები არ ჩაიწერება)_")
            mash_template = pd.DataFrame({
                "საფეხური #": [1, 2, 3, 4, 5],
                "ტემპ (°C)": [0.0] * 5,
                "ხანგრძლივობა (წთ)": [0] * 5,
                "შენიშვნა": [""] * 5,
            })
            mash_edit = st.data_editor(
                mash_template, num_rows="fixed", use_container_width=True,
                key="mash_steps_editor", disabled=locked,
            )
            if st.button("💾 საფეხურების შენახვა", disabled=locked):
                # Phase D.2: block duplicate temperatures — both against steps
                # already in BREW_STEPS and within the 5-row batch itself.
                existing_mash = get_df("BREW_STEPS")
                existing_temps = {}  # temp(°C) -> step Item name already saved
                if (not existing_mash.empty
                        and {"Brew_ID", "Category", "Qty", "Item"} <= set(existing_mash.columns)):
                    em = existing_mash[(existing_mash["Brew_ID"] == bid)
                                       & (existing_mash["Category"] == "Mash_Step")]
                    for _, er in em.iterrows():
                        existing_temps[_num(er.get("Qty"))] = str(er.get("Item", ""))

                to_save, batch_temps, dup_err = [], {}, None
                for _, r in mash_edit.iterrows():
                    temp_v = r["ტემპ (°C)"]
                    dur_v = r["ხანგრძლივობა (წთ)"]
                    if (temp_v in (None, "", 0, 0.0)) and (dur_v in (None, "", 0)):
                        continue  # ცარიელი row — იგნორი
                    t = _num(temp_v)
                    step_no = int(r["საფეხური #"])
                    if t in existing_temps:
                        dup_err = (f"ეს ტემპერატურა ({t:g}°C) უკვე დამატებულია "
                                   f"საფეხურ „{existing_temps[t]}“-ში.")
                        break
                    if t in batch_temps:
                        dup_err = (f"ეს ტემპერატურა ({t:g}°C) გამეორებულია — "
                                   f"საფეხური {batch_temps[t]} და {step_no}.")
                        break
                    batch_temps[t] = step_no
                    to_save.append((step_no, temp_v, dur_v, r["შენიშვნა"]))

                if dup_err:
                    st.error(dup_err)
                elif not to_save:
                    st.success("შესავსები row ვერ მოიძებნა.")
                else:
                    for step_no, temp_v, dur_v, note in to_save:
                        append_row("BREW_STEPS", {
                            "Brew_ID": bid, "Stage": "Mash", "Category": "Mash_Step",
                            "Item": f"საფეხური {step_no}",
                            "Qty": temp_v, "Unit": "°C",
                            "Timing": f"{int(dur_v)} წთ", "Justification": note,
                        })
                    st.success(f"{len(to_save)} საფეხური შენახულია.")
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
                c1, c2 = st.columns(2)
                malt_qty = c1.number_input(
                    f"რაოდენობა (kg) — მარაგში {malt_cap:g} kg" if malt_cap is not None
                    else "რაოდენობა (kg)",
                    min_value=0.0,
                    # historical entries may exceed today's stock — no cap then
                    max_value=malt_cap if (malt_cap and malt_cap > 0 and not hist_mode) else None,
                    key="malt_qty", disabled=locked)
                malt_just = c2.text_input("დასაბუთება", key="malt_just", disabled=locked)
                malt_deduct = convert_to_inventory_unit(malt_qty, "kg", malt_inv_unit)
                # Phase E: same Brew_ID+Category+Item already added?
                malt_steps = get_df("BREW_STEPS")
                malt_existing = (malt_steps[(malt_steps["Brew_ID"] == bid)
                                            & (malt_steps["Category"] == "Malt")
                                            & (malt_steps["Item"] == malt_name)]
                                 if not malt_steps.empty
                                 and {"Brew_ID", "Category", "Item"} <= set(malt_steps.columns)
                                 else pd.DataFrame())
                if hist_mode:
                    st.caption("📜 ისტორიული რეჟიმი — მარაგიდან არ ჩამოეჭრება.")
                elif malt_deduct:
                    if not malt_existing.empty:
                        prev = _num(malt_existing.iloc[0].get("Qty"))
                        st.caption(f"უკვე დამატებულია {prev:g} kg — დაემატება ჯამში "
                                   f"{prev + malt_qty:g} kg. მარაგიდან ჩამოეჭრება: "
                                   f"{malt_deduct:g} {malt_inv_unit}")
                    else:
                        st.caption(f"მარაგიდან ჩამოეჭრება: {malt_deduct:g} {malt_inv_unit}")
                if st.button("დამატება (ალაო)", disabled=locked) and malt_qty > 0:
                    if not hist_mode and malt_deduct is None:
                        st.error("ერთეულები შეუთავსებელია — ჩანაწერი არ შენახულა, "
                                 "ჩამოჭრა არ მომხდარა.")
                    elif not hist_mode and malt_deduct > malt_stock:
                        st.error(f"მარაგშია მხოლოდ {malt_stock:g} {malt_inv_unit}, "
                                 f"მოთხოვნილია {malt_deduct:g} {malt_inv_unit}.")
                    else:
                        if not malt_existing.empty:
                            new_total = round(_num(malt_existing.iloc[0].get("Qty")) + malt_qty, 4)
                            update_step_qty(bid, "Malt", malt_name, new_total)
                        else:
                            append_row("BREW_STEPS", {
                                "Brew_ID": bid, "Stage": "Mash", "Category": "Malt",
                                "Item": malt_name, "Qty": malt_qty, "Unit": "kg",
                                "Justification": malt_just,
                            })
                        if not hist_mode:
                            update_cell_by_key("INVENTORY", "Item", malt_name,
                                               "Current_Qty", round(malt_stock - malt_deduct, 4))
                        st.rerun()

            steps_df = get_df("BREW_STEPS")
            if not steps_df.empty:
                mash_view = steps_df[(steps_df["Brew_ID"] == bid) & (steps_df["Stage"] == "Mash")]
                st.dataframe(mash_view, use_container_width=True)

        # ---- BOIL / HOP TAB ----
        if st.session_state.brew_tab == "🔥 დუღილი (boil/hop)":
            locked = lock_gate(row.get("FG_Date"), f"edit_ovr_boil_{bid}") if fg_done else False
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
                c1, c2, c3 = st.columns(3)
                hop_qty = c1.number_input(
                    f"რაოდენობა (g) — მარაგში {hop_cap:g} g" if hop_cap is not None
                    else "რაოდენობა (g)",
                    min_value=0.0,
                    # historical entries may exceed today's stock — no cap then
                    max_value=hop_cap if (hop_cap and hop_cap > 0 and not hist_mode) else None,
                    key="hop_qty", disabled=locked)
                hop_timing_sel = c2.selectbox(
                    "დრო", ["60წთ", "30წთ", "15წთ", "5წთ", "0", "Whirlpool", "სხვა"],
                    key="hop_timing", disabled=locked)
                c3.text_input("AA% (INVENTORY-დან)", value=hop_aa, disabled=True, key="hop_aa_display")
                if hop_timing_sel == "სხვა":
                    hop_timing = st.text_input("დრო — ხელით (მაგ. 45წთ, 20წთ)",
                                               key="hop_timing_custom", disabled=locked).strip()
                else:
                    hop_timing = hop_timing_sel
                hop_just = st.text_input("დასაბუთება", key="hop_just", disabled=locked)
                hop_deduct = convert_to_inventory_unit(hop_qty, "g", hop_inv_unit)
                # Phase E: true duplicate for hops = same Brew_ID+Category+Item
                # AND Timing. Same hop at another timing is a separate addition
                # (e.g. Tradition@30წთ vs Tradition@0წთ must stay two rows).
                hop_steps = get_df("BREW_STEPS")
                hop_existing = (hop_steps[(hop_steps["Brew_ID"] == bid)
                                          & (hop_steps["Category"] == "Hop")
                                          & (hop_steps["Item"] == hop_name)
                                          & (hop_steps["Timing"].astype(str) == str(hop_timing))]
                                if not hop_steps.empty
                                and {"Brew_ID", "Category", "Item", "Timing"} <= set(hop_steps.columns)
                                else pd.DataFrame())
                if hist_mode:
                    st.caption("📜 ისტორიული რეჟიმი — მარაგიდან არ ჩამოეჭრება.")
                elif hop_deduct:
                    if not hop_existing.empty:
                        prev = _num(hop_existing.iloc[0].get("Qty"))
                        st.caption(f"უკვე დამატებულია {prev:g} g ({hop_timing}) — "
                                   f"დაემატება ჯამში {prev + hop_qty:g} g. "
                                   f"მარაგიდან ჩამოეჭრება: {hop_deduct:g} {hop_inv_unit}")
                    else:
                        st.caption(f"მარაგიდან ჩამოეჭრება: {hop_deduct:g} {hop_inv_unit}")
                if st.button("დამატება (ჰოპი)", disabled=locked) and hop_qty > 0:
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
                                            timing=hop_timing)
                        else:
                            append_row("BREW_STEPS", {
                                "Brew_ID": bid, "Stage": "Boil", "Category": "Hop",
                                "Item": hop_name, "Qty": hop_qty, "Unit": "g",
                                "Timing": hop_timing, "AA_%": hop_aa, "Justification": hop_just,
                            })
                        if not hist_mode:
                            update_cell_by_key("INVENTORY", "Item", hop_name,
                                               "Current_Qty", round(hop_stock - hop_deduct, 4))
                        st.rerun()

            c1, c2 = st.columns(2)
            pre_boil_vol = c1.number_input("Pre-Boil მოცულობა (L)", disabled=locked)
            pre_boil_p = c2.number_input("Pre-Boil Gravity (°P)", disabled=locked)
            c3, c4 = st.columns(2)
            post_boil_vol = c3.number_input("Post-Boil მოცულობა (L)", disabled=locked)
            actual_og = c4.number_input("Actual OG (°P)", disabled=locked)
            if st.button("💾 boil მონაცემის შენახვა header-ში", disabled=locked):
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
                st.success("შენახულია.")
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
                st.success(f"საფუარი შენახულია: {yeast_name} ({form_val}"
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
                if add_g and date_error:
                    st.error("თარიღი ხარშვის დაწყებამდეა — ჩანაწერი არ დაემატა.")
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
