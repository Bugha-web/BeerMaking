import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime, date
import uuid

st.set_page_config(page_title="BUGHASHVILI Brew Journal", layout="wide")

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
    data = ws(name).get_all_records()
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
page = st.sidebar.radio("გვერდი", ["📦 Inventory", "🍺 ხარშვა"])

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
        notes = st.text_input("შენიშვნა")
        submitted = st.form_submit_button("დამატება")
        if submitted and item:
            append_row("INVENTORY", {
                "Item": item, "Category": cat, "Unit": unit, "Current_Qty": qty,
                "AA_%_if_hop": aa, "Low_Threshold": low,
                "Received_Date": str(date.today()), "Notes": notes,
            })
            st.success(f"{item} დაემატა.")
            st.rerun()

# ============================================================
# PAGE 2 — BREW
# ============================================================
else:
    st.title("🍺 ხარშვის ჟურნალი")

    headers_df = get_df("BREW_HEADER")
    ids = headers_df["Brew_ID"].tolist() if not headers_df.empty else []
    choice = st.selectbox("აირჩიე ხარშვა ან შექმენი ახალი", ["➕ ახალი ხარშვა"] + ids)

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
                append_row("BREW_HEADER", {
                    "Brew_ID": bid, "Date": str(b_date), "Beer_Style": style,
                    "Fermenter": ferm, "Target_Vol_L": target_vol, "Water_uS": water_us,
                    "Target_OG_P": target_og, "Target_FG_P": target_fg,
                })
                st.success(f"ხარშვა შეიქმნა: {bid}. აირჩიე ის ზემოთა სიიდან რომ დეტალები შეავსო.")
                st.rerun()

    # ---------- EXISTING BREW: TABS ----------
    else:
        bid = choice
        row = headers_df[headers_df["Brew_ID"] == bid].iloc[0]
        fg_done = str(row.get("Actual_FG_P", "")) not in ("", "nan", "None")

        st.subheader(f"{bid} — {row.get('Beer_Style','')}")
        if fg_done:
            st.info("✅ FG მიღწეულია — ეს ხარშვა დახურულია. ცვლილება კვლავ შესაძლებელია, მაგრამ საჭირო აღარ არის.")

        BREW_TABS = ["💧 წყალი", "🌾 მეშინგი", "🔥 დუღილი (boil/hop)", "🧪 ფერმენტაცია/Gravity"]
        if "brew_tab" not in st.session_state:
            st.session_state.brew_tab = "💧 წყალი"
        # tab-like look for the horizontal radio; scoped via the widget key
        # class (.st-key-*) so the sidebar radio is unaffected
        st.markdown("""
            <style>
            .st-key-brew_tab_radio div[role="radiogroup"] label {
                border: 1px solid rgba(128,128,128,.4); border-bottom: none;
                border-radius: 8px 8px 0 0; padding: 4px 14px; margin-right: 4px;
            }
            .st-key-brew_tab_radio div[role="radiogroup"] label > div:first-child {
                display: none;
            }
            </style>
        """, unsafe_allow_html=True)
        st.session_state.brew_tab = st.radio(
            "ტაბი", BREW_TABS,
            horizontal=True, label_visibility="collapsed", key="brew_tab_radio",
            index=BREW_TABS.index(st.session_state.brew_tab)
        )

        # ---- WATER TAB ----
        if st.session_state.brew_tab == "💧 წყალი":
            profiles_df = get_df("WATER_PROFILES")
            profile_names = sorted(profiles_df["Profile_Name"].unique()) if not profiles_df.empty else []
            chosen_profile = st.selectbox("წყლის პროფილი (თუ შენახული გაქვს)",
                                           ["— ხელით შევსება —"] + profile_names)

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
            m_vol = m1.number_input("მოცულობა (L)", min_value=0.0, key="m_vol")
            m_gyp = m2.number_input("Gypsum (g)", value=float(defaults["Mash"]["gyp"] or 0), key="m_gyp")
            m_cacl = m3.number_input("CaCl2 (g)", value=float(defaults["Mash"]["cacl2"] or 0), key="m_cacl")
            m_lac = m4.number_input("Lactic (ml)", value=float(defaults["Mash"]["lac"] or 0), key="m_lac")
            m_ph = m5.number_input("Target pH", value=float(defaults["Mash"]["ph"] or 5.3), key="m_ph")

            st.markdown("**Sparge წყალი**")
            s1, s2, s3, s4, s5 = st.columns(5)
            s_vol = s1.number_input("მოცულობა (L)", value=float(defaults["Sparge"]["vol"] or 1000), key="s_vol")
            s_gyp = s2.number_input("Gypsum (g)", value=float(defaults["Sparge"]["gyp"] or 0), key="s_gyp")
            s_cacl = s3.number_input("CaCl2 (g)", value=float(defaults["Sparge"]["cacl2"] or 0), key="s_cacl")
            s_lac = s4.number_input("Lactic (ml)", value=float(defaults["Sparge"]["lac"] or 0), key="s_lac")
            s_ph = s5.number_input("Target pH", value=float(defaults["Sparge"]["ph"] or 5.7), key="s_ph")

            if st.button("💾 წყლის მონაცემის შენახვა"):
                append_row("WATER_TREATMENT", {"Brew_ID": bid, "Water_Stream": "Mash",
                    "Volume_L": m_vol, "Gypsum_g": m_gyp, "CaCl2_g": m_cacl,
                    "Lactic_ml": m_lac, "Target_pH": m_ph})
                append_row("WATER_TREATMENT", {"Brew_ID": bid, "Water_Stream": "Sparge",
                    "Volume_L": s_vol, "Gypsum_g": s_gyp, "CaCl2_g": s_cacl,
                    "Lactic_ml": s_lac, "Target_pH": s_ph})
                st.success("წყლის მონაცემი შენახულია.")

            st.caption("იმ პროფილს ხედავ პირველად? — შეინახე ახლანდელი მნიშვნელობები, რომ მომდევნო ჯერზე აღარ გჭირდეს ხელით.")
            new_profile_name = st.text_input("შეინახე ეს, როგორც ახალი პროფილი (სახელი)")
            if st.button("💾 პროფილად შენახვა") and new_profile_name:
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
                key="mash_steps_editor",
            )
            if st.button("💾 საფეხურების შენახვა"):
                saved = 0
                for _, r in mash_edit.iterrows():
                    temp_v = r["ტემპ (°C)"]
                    dur_v = r["ხანგრძლივობა (წთ)"]
                    if (temp_v in (None, "", 0, 0.0)) and (dur_v in (None, "", 0)):
                        continue  # ცარიელი row — იგნორი
                    append_row("BREW_STEPS", {
                        "Brew_ID": bid, "Stage": "Mash", "Category": "Mash_Step",
                        "Item": f"საფეხური {int(r['საფეხური #'])}",
                        "Qty": temp_v, "Unit": "°C",
                        "Timing": f"{int(dur_v)} წთ", "Justification": r["შენიშვნა"],
                    })
                    saved += 1
                st.success(f"{saved} საფეხური შენახულია." if saved else "შესავსები row ვერ მოიძებნა.")
                if saved:
                    st.rerun()

            st.markdown("**ალაოს ჩამონათვალი (grain bill)**")
            st.caption("[Certain] ვალიდაცია მხოლოდ — მარაგზე მეტს ვერ აირჩევ. "
                       "INVENTORY-დან ავტომატური ჩამოჭრა განზრახ არ ხდება.")
            inv_df = get_df("INVENTORY")
            malt_opts = (sorted(inv_df[inv_df["Category"] == "Malt"]["Item"].tolist())
                         if not inv_df.empty and "Category" in inv_df.columns else [])
            if not malt_opts:
                st.info("INVENTORY-ში Malt კატეგორიის ჩანაწერი არ არის — ჯერ დაამატე მარაგში.")
            else:
                malt_name = st.selectbox("ალაო", malt_opts, key="malt_sel")
                malt_max = _num(inv_df[inv_df["Item"] == malt_name].iloc[0].get("Current_Qty", 0))
                c1, c2 = st.columns(2)
                malt_qty = c1.number_input(
                    f"რაოდენობა (kg) — მარაგში {malt_max:g}", min_value=0.0,
                    max_value=malt_max if malt_max > 0 else None, key="malt_qty")
                malt_just = c2.text_input("დასაბუთება", key="malt_just")
                if st.button("დამატება (ალაო)") and malt_qty > 0:
                    append_row("BREW_STEPS", {
                        "Brew_ID": bid, "Stage": "Mash", "Category": "Malt",
                        "Item": malt_name, "Qty": malt_qty, "Unit": "kg",
                        "Justification": malt_just,
                    })
                    st.rerun()

            steps_df = get_df("BREW_STEPS")
            if not steps_df.empty:
                mash_view = steps_df[(steps_df["Brew_ID"] == bid) & (steps_df["Stage"] == "Mash")]
                st.dataframe(mash_view, use_container_width=True)

        # ---- BOIL / HOP TAB ----
        if st.session_state.brew_tab == "🔥 დუღილი (boil/hop)":
            st.markdown("**ჰოპის დამატება**")
            st.caption("[Certain] ვალიდაცია მხოლოდ — მარაგზე მეტს ვერ აირჩევ. "
                       "AA% ავტომატურად ივსება INVENTORY-დან. ჩამოჭრა არ ხდება.")
            inv_hop_df = get_df("INVENTORY")
            hop_opts = (sorted(inv_hop_df[inv_hop_df["Category"] == "Hop"]["Item"].tolist())
                        if not inv_hop_df.empty and "Category" in inv_hop_df.columns else [])
            if not hop_opts:
                st.info("INVENTORY-ში Hop კატეგორიის ჩანაწერი არ არის — ჯერ დაამატე მარაგში.")
            else:
                hop_name = st.selectbox("ჰოპი", hop_opts, key="hop_sel")
                hop_inv_row = inv_hop_df[inv_hop_df["Item"] == hop_name].iloc[0]
                hop_max = _num(hop_inv_row.get("Current_Qty", 0))
                hop_aa = str(hop_inv_row.get("AA_%_if_hop", "") or "")
                c1, c2, c3 = st.columns(3)
                hop_qty = c1.number_input(
                    f"რაოდენობა (g) — მარაგში {hop_max:g}", min_value=0.0,
                    max_value=hop_max if hop_max > 0 else None, key="hop_qty")
                hop_timing = c2.selectbox("დრო", ["60წთ", "30წთ", "15წთ", "5წთ", "0", "Whirlpool"], key="hop_timing")
                c3.text_input("AA% (INVENTORY-დან)", value=hop_aa, disabled=True, key="hop_aa_display")
                hop_just = st.text_input("დასაბუთება", key="hop_just")
                if st.button("დამატება (ჰოპი)") and hop_qty > 0:
                    append_row("BREW_STEPS", {
                        "Brew_ID": bid, "Stage": "Boil", "Category": "Hop",
                        "Item": hop_name, "Qty": hop_qty, "Unit": "g",
                        "Timing": hop_timing, "AA_%": hop_aa, "Justification": hop_just,
                    })
                    st.rerun()

            c1, c2 = st.columns(2)
            pre_boil_vol = c1.number_input("Pre-Boil მოცულობა (L)")
            pre_boil_p = c2.number_input("Pre-Boil Gravity (°P)")
            c3, c4 = st.columns(2)
            post_boil_vol = c3.number_input("Post-Boil მოცულობა (L)")
            actual_og = c4.number_input("Actual OG (°P)")
            if st.button("💾 boil მონაცემის შენახვა header-ში"):
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
            st.markdown("**ყოველდღიური Gravity-ის ჩაწერა**")
            with st.form("gravity_add", clear_on_submit=True):
                c1, c2, c3, c4 = st.columns(4)
                g_date = c1.date_input("თარიღი", value=date.today())
                g_day = c2.number_input("დღე #", min_value=0, step=1)
                g_raw = c3.number_input("Gravity (°P) — ჰიდრომეტრიდან, დაუკორექტირებელი")
                g_temp = c4.number_input("ტემპერატურა (°C)", value=20.0)
                is_final = st.checkbox("ეს არის საბოლოო (FG) გაზომვა")
                add_g = st.form_submit_button("დამატება")
                if add_g:
                    ref_temp = _num(get_setting("Hydrometer_Ref_Temp_C", 20), 20)
                    corrected = correct_gravity_plato(g_raw, g_temp, ref_temp)
                    append_row("BREW_GRAVITY_LOG", {
                        "Brew_ID": bid, "Date": str(g_date), "Day_#": g_day,
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
                                    value=str(row.get("Outcome_Note", "")))
            if st.button("💾 შედეგის შენახვა"):
                update_cell_by_key("BREW_HEADER", "Brew_ID", bid, "Outcome_Note", outcome)
                st.success("შენახულია.")
