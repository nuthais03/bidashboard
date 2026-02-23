import io
import time
from datetime import datetime

import numpy as np
import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.io as pio

import bcrypt
import streamlit_authenticator as stauth

# PDF
PDF_AVAILABLE = True
try:
    from reportlab.lib.pagesizes import A4
    from reportlab.pdfgen import canvas
    from reportlab.lib.units import cm
except ModuleNotFoundError:
    PDF_AVAILABLE = False

# Google Sheets
GSHEETS_AVAILABLE = True
try:
    import gspread
    from google.oauth2.service_account import Credentials
except ModuleNotFoundError:
    GSHEETS_AVAILABLE = False


# -----------------------------------
# Branding
# -----------------------------------
OWT_LOGO_URL = "https://owtgroupltd.co.uk/assets/images/OWT%20Group%20Logo-TransparentBackground1.png"
APP_TITLE = "OWT Marketing Performance Dashboard"
APP_TAGLINE = "Private performance dashboard — client-ready SaaS version."
BRAND_ACCENT = "#3B82F6"

MONTH_ORDER = [
    "January","February","March","April","May","June",
    "July","August","September","October","November","December"
]

FINAL_COL_ORDER = [
    "month","brand","destination","impressions","cpl","spent_gbp","leads","converted_leads","conversion_rate"
]

MANUAL_UI_COLS = ["Month","Brand","Destination","Impressions","Spent (GBP)","Leads","Converted Leads"]

SESSION_TIMEOUT_SECONDS = 20 * 60  # 20 min idle


# -----------------------------------
# Utilities: CSS / Theme / White-label
# -----------------------------------
def hide_streamlit_branding():
    st.markdown("""
    <style>
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    [data-testid="stToolbar"] {display:none;}
    [data-testid="stDecoration"] {display:none;}
    </style>
    """, unsafe_allow_html=True)


def apply_theme(theme_mode: str) -> None:
    if theme_mode == "Light":
        pio.templates.default = "plotly_white"
        st.markdown(
            f"""
            <style>
              .block-container {{ padding-top: 1.0rem; padding-bottom: 2rem; }}
              section[data-testid="stSidebar"] {{ background-color: #ffffff; }}
              .stApp {{ background-color: #f5f7fb; color: #111111; }}
              h1,h2,h3 {{ letter-spacing: 0.2px; color:#111111; }}
              [data-testid="stMetricValue"] {{ font-size: 1.6rem; color:#111111; }}
              .card {{
                border:1px solid #e5e7eb; border-radius:14px; padding:14px; background:#ffffff;
              }}
              .muted {{ opacity:0.75; color:#111111; }}
              .header-bar {{
                border: 1px solid rgba(17,17,17,0.10);
                background: rgba(255,255,255,0.75);
                border-radius: 14px;
                padding: 10px 12px;
              }}
              .pill {{
                border: 1px solid rgba(17,17,17,0.16);
                border-radius: 999px;
                padding: 6px 10px;
                font-size: 12px;
                background: rgba(17,17,17,0.03);
                color:#111111;
                display:inline-flex;
                align-items:center;
                gap:6px;
              }}
              .dot {{
                height: 8px; width: 8px; border-radius: 50%;
                background: {BRAND_ACCENT}; display: inline-block;
              }}
            </style>
            """,
            unsafe_allow_html=True,
        )
    else:
        pio.templates.default = "plotly_dark"
        st.markdown(
            f"""
            <style>
              .block-container {{ padding-top: 1.0rem; padding-bottom: 2rem; }}
              section[data-testid="stSidebar"] {{ background-color: #0b0f17; }}
              .stApp {{ background-color: #0b0f17; color: #ffffff; }}
              h1,h2,h3 {{ letter-spacing: 0.2px; }}
              [data-testid="stMetricValue"] {{ font-size: 1.6rem; }}
              .card {{
                border:1px solid rgba(255,255,255,0.10);
                border-radius:14px; padding:14px; background:rgba(255,255,255,0.02);
              }}
              .muted {{ opacity:0.85; }}
              .header-bar {{
                border: 1px solid rgba(255,255,255,0.10);
                background: rgba(255,255,255,0.03);
                border-radius: 14px;
                padding: 10px 12px;
              }}
              .pill {{
                border: 1px solid rgba(255,255,255,0.14);
                border-radius: 999px;
                padding: 6px 10px;
                font-size: 12px;
                background: rgba(255,255,255,0.03);
                display:inline-flex;
                align-items:center;
                gap:6px;
              }}
              .dot {{
                height: 8px; width: 8px; border-radius: 50%;
                background: {BRAND_ACCENT}; display: inline-block;
              }}
            </style>
            """,
            unsafe_allow_html=True,
        )


# -----------------------------------
# Google Sheets helpers
# -----------------------------------
def _get_gs_client():
    if not GSHEETS_AVAILABLE:
        raise RuntimeError("Missing Google Sheets libs. Add gspread + google-auth.")
    if "gcp_service_account" not in st.secrets or "gsheet_id" not in st.secrets:
        raise RuntimeError("Missing secrets: gcp_service_account + gsheet_id")
    creds_info = dict(st.secrets["gcp_service_account"])
    scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    creds = Credentials.from_service_account_info(creds_info, scopes=scopes)
    return gspread.authorize(creds)

def _open_sheet():
    client = _get_gs_client()
    return client.open_by_key(st.secrets["gsheet_id"])

def ws_data():
    sh = _open_sheet()
    return sh.worksheet(st.secrets.get("gsheet_tab", "Data"))

def ws_users():
    sh = _open_sheet()
    return sh.worksheet(st.secrets.get("gsheet_users_tab", "Users"))

def ws_audit():
    sh = _open_sheet()
    return sh.worksheet(st.secrets.get("gsheet_audit_tab", "Audit"))


def audit_log(action: str, meta: dict | None = None):
    meta = meta or {}
    try:
        u = st.session_state.get("auth_user", {})
        row = [
            datetime.utcnow().isoformat(),
            str(u.get("username", "")),
            str(u.get("role", "")),
            str(u.get("client_id", "")),
            action,
            str(meta),
        ]
        ws_audit().append_row(row, value_input_option="RAW")
    except Exception:
        # Don't break app if audit fails
        pass


# -----------------------------------
# Users + Auth (from Google Sheet)
# -----------------------------------
def load_users_from_sheet() -> pd.DataFrame:
    ws = ws_users()
    values = ws.get_all_values()
    if not values or len(values) < 2:
        return pd.DataFrame(columns=["username","name","email","password_hash","role","client_id","active"])
    header = values[0]
    rows = values[1:]
    df = pd.DataFrame(rows, columns=header)
    # normalize
    for col in ["username","name","email","password_hash","role","client_id","active"]:
        if col not in df.columns:
            df[col] = ""
    df = df[["username","name","email","password_hash","role","client_id","active"]].copy()
    df["active"] = df["active"].astype(str).str.upper().isin(["TRUE","1","YES","Y"])
    df["username"] = df["username"].astype(str).str.strip()
    df["role"] = df["role"].astype(str).str.strip().replace("", "client")
    df["client_id"] = df["client_id"].astype(str).str.strip().replace("", "DEFAULT")
    return df[df["username"] != ""]

def credentials_from_users_df(users_df: pd.DataFrame) -> dict:
    creds = {"usernames": {}}
    for _, r in users_df.iterrows():
        if not bool(r["active"]):
            continue
        creds["usernames"][r["username"]] = {
            "name": r["name"],
            "email": r["email"],
            "password": r["password_hash"],
            "role": r["role"],
            "client_id": r["client_id"],
        }
    return creds


def require_login() -> tuple[stauth.Authenticate, dict]:
    users_df = load_users_from_sheet()
    creds = credentials_from_users_df(users_df)

    authenticator = stauth.Authenticate(
        credentials=creds,
        cookie_name=st.secrets["auth"]["cookie_name"],
        key=st.secrets["auth"]["cookie_key"],
        cookie_expiry_days=int(st.secrets["auth"]["cookie_expiry_days"]),
    )

    name, auth_status, username = authenticator.login(location="main")

    if auth_status is False:
        st.error("Invalid username or password")
        st.stop()
    if auth_status is None:
        st.markdown(f"## {APP_TITLE}")
        st.info("Please log in to continue.")
        st.stop()

    # Logged in
    user = creds["usernames"][username]
    st.session_state["auth_user"] = {
        "username": username,
        "name": user["name"],
        "email": user["email"],
        "role": user.get("role", "client"),
        "client_id": user.get("client_id", "DEFAULT"),
    }

    # Log login once per session
    if not st.session_state.get("_login_logged", False):
        audit_log("login_success", {})
        st.session_state["_login_logged"] = True

    # idle timer
    st.session_state["last_active"] = time.time()
    return authenticator, st.session_state["auth_user"]


def enforce_idle_timeout(authenticator: stauth.Authenticate):
    now = time.time()
    last = st.session_state.get("last_active", now)
    if now - last > SESSION_TIMEOUT_SECONDS:
        audit_log("session_timeout", {})
        st.warning("Session timed out due to inactivity. Please log in again.")
        authenticator.logout("Login again", location="main")
        st.session_state.clear()
        st.stop()
    st.session_state["last_active"] = now


# -----------------------------------
# Data helpers
# -----------------------------------
def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = df.columns.str.strip()

    rename_map = {
        "Month": "month",
        "Brand": "brand",
        "Destination": "destination",
        "Impressions": "impressions",
        "Spent (GBP)": "spent_gbp",
        "Spent": "spent_gbp",
        "Spend": "spent_gbp",
        "Leads": "leads",
        "Converted Leads": "converted_leads",
        "CPL": "cpl",
        "Conversion Rate": "conversion_rate",
        "Client ID": "client_id",
        "client_id": "client_id",
    }

    lower_to_actual = {c.lower(): c for c in df.columns}
    for k, v in list(rename_map.items()):
        if k not in df.columns and k.lower() in lower_to_actual:
            rename_map[lower_to_actual[k.lower()]] = v

    df = df.rename(columns={k: v for k, v in rename_map.items() if k in df.columns})

    required = {"month","brand","destination","leads","spent_gbp"}
    missing = required - set(df.columns)
    if missing:
        raise ValueError("Missing required columns: " + ", ".join(sorted(missing)) +
                         " | Required: Month, Brand, Destination, Leads, Spent (GBP)")

    if "impressions" not in df.columns:
        df["impressions"] = 0
    if "converted_leads" not in df.columns:
        df["converted_leads"] = 0
    if "client_id" not in df.columns:
        df["client_id"] = "DEFAULT"

    df["month"] = df["month"].astype(str).str.strip()
    df["brand"] = df["brand"].astype(str).str.strip()
    df["destination"] = df["destination"].astype(str).str.strip()
    df["client_id"] = df["client_id"].astype(str).str.strip()

    df["impressions"] = pd.to_numeric(df["impressions"], errors="coerce").fillna(0).astype(int)
    df["spent_gbp"] = pd.to_numeric(df["spent_gbp"], errors="coerce").fillna(0.0).astype(float)
    df["leads"] = pd.to_numeric(df["leads"], errors="coerce").fillna(0).astype(int)
    df["converted_leads"] = pd.to_numeric(df["converted_leads"], errors="coerce").fillna(0).astype(int)

    if df["month"].isin(MONTH_ORDER).any():
        df["month"] = pd.Categorical(df["month"], categories=MONTH_ORDER, ordered=True)

    return df

def compute_metrics(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    leads_safe = out["leads"].replace(0, np.nan)
    out["cpl"] = (out["spent_gbp"] / leads_safe).fillna(0.0)
    out["conversion_rate"] = (out["converted_leads"] / leads_safe).fillna(0.0)
    return out

def apply_overrides(base_df: pd.DataFrame, overrides: dict) -> pd.DataFrame:
    df = base_df.copy()
    if "row_id" not in df.columns:
        df["row_id"] = df["month"].astype(str) + "||" + df["brand"].astype(str) + "||" + df["destination"].astype(str)

    imp_map = overrides.get("impressions", {})
    conv_map = overrides.get("converted_leads", {})

    if imp_map:
        df["impressions"] = df["row_id"].map(imp_map).combine_first(df["impressions"]).astype(int)
    if conv_map:
        df["converted_leads"] = df["row_id"].map(conv_map).combine_first(df["converted_leads"]).astype(int)

    return df


def build_pdf_report(filters: dict, d: pd.DataFrame, user: dict) -> bytes:
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    _, h = A4

    c.setFont("Helvetica-Bold", 16)
    c.drawString(2 * cm, h - 2 * cm, "OWT — Leads Performance Report")

    c.setFont("Helvetica", 10)
    c.drawString(2 * cm, h - 2.8 * cm, f"Client: {user.get('client_id','')} | User: {user.get('username','')}")
    c.drawString(
        2 * cm,
        h - 3.4 * cm,
        f"Month: {filters.get('month','All')} | Brand: {filters.get('brand','All')} | Destination: {filters.get('destination','All')}",
    )

    total_spend = float(d["spent_gbp"].sum())
    total_leads = int(d["leads"].sum())
    total_impr = int(d["impressions"].sum())
    total_conv = int(d["converted_leads"].sum())
    overall_cpl = (total_spend / total_leads) if total_leads > 0 else 0.0
    overall_cr = (total_conv / total_leads) if total_leads > 0 else 0.0

    y = h - 5.0 * cm
    c.setFont("Helvetica-Bold", 12)
    c.drawString(2 * cm, y, "Summary KPIs")
    y -= 0.8 * cm

    c.setFont("Helvetica", 11)
    c.drawString(2 * cm, y, f"Total Spend: £{total_spend:,.2f}")
    c.drawString(9 * cm, y, f"Total Leads: {total_leads:,}")
    y -= 0.6 * cm
    c.drawString(2 * cm, y, f"Total Impressions: {total_impr:,}")
    c.drawString(9 * cm, y, f"Overall CPL: £{overall_cpl:,.2f}")
    y -= 0.6 * cm
    c.drawString(2 * cm, y, f"Total Converted Leads: {total_conv:,}")
    c.drawString(9 * cm, y, f"Overall Conversion Rate: {overall_cr*100:,.2f}%")

    c.showPage()
    c.save()
    buf.seek(0)
    return buf.read()


# -----------------------------------
# App start
# -----------------------------------
st.set_page_config(page_title=APP_TITLE, layout="wide")
hide_streamlit_branding()

if "theme_mode" not in st.session_state:
    st.session_state.theme_mode = "Dark"
apply_theme(st.session_state.theme_mode)

# Session state
if "mode" not in st.session_state:
    st.session_state.mode = None
if "overrides" not in st.session_state:
    st.session_state.overrides = {"impressions": {}, "converted_leads": {}}
if "manual_data" not in st.session_state:
    st.session_state.manual_data = pd.DataFrame(columns=MANUAL_UI_COLS)

# Auth
authenticator, user = require_login()
enforce_idle_timeout(authenticator)

# Sidebar SaaS
st.sidebar.image(OWT_LOGO_URL, width=160)
st.sidebar.markdown(f"**{user['name']}**  \n`{user['role']}`  \nClient: **{user['client_id']}**")
if st.sidebar.button("Logout", use_container_width=True):
    audit_log("logout", {})
    authenticator.logout("Logout", location="main")
    st.session_state.clear()
    st.rerun()

st.sidebar.markdown("---")

# App navigation (SaaS style)
menu = st.sidebar.radio("Menu", ["Dashboard", "Messages", "Settings"], index=0)
if menu == "Settings" and user["role"] != "admin":
    st.error("Access denied (admin only).")
    st.stop()

# Header bar + theme toggle
st.markdown('<div class="header-bar">', unsafe_allow_html=True)
h1, h2, h3, h4 = st.columns([2, 6, 2, 2])

with h1:
    st.image(OWT_LOGO_URL, width=150)
with h2:
    st.markdown(f"### {APP_TITLE}")
    st.caption(APP_TAGLINE)
with h3:
    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown(f'<span class="pill"><span class="dot"></span>Client Access</span>', unsafe_allow_html=True)
with h4:
    st.markdown("<br>", unsafe_allow_html=True)
    light = st.toggle("Light mode", value=(st.session_state.theme_mode == "Light"))
    new_mode = "Light" if light else "Dark"
    if new_mode != st.session_state.theme_mode:
        st.session_state.theme_mode = new_mode
        st.rerun()

st.markdown("</div>", unsafe_allow_html=True)
st.write("")


# -----------------------------------
# SETTINGS PAGE (Admin Only)
# -----------------------------------
def settings_page():
    st.subheader("Settings (Admin)")
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown("### User Management")
    st.markdown('<div class="muted">Users are stored in Google Sheet tab: <b>Users</b>. Passwords are stored as <b>bcrypt hashes</b>.</div>', unsafe_allow_html=True)

    # Load current users
    try:
        users_df = load_users_from_sheet()
    except Exception as e:
        st.error(f"Could not load Users sheet: {e}")
        st.markdown("</div>", unsafe_allow_html=True)
        return

    st.dataframe(users_df, use_container_width=True, hide_index=True)

    st.markdown("### Add / Update User")
    c1, c2 = st.columns(2)
    with c1:
        username = st.text_input("username (unique)", placeholder="client3")
        name = st.text_input("name", placeholder="Client 3")
        email = st.text_input("email", placeholder="client3@client.com")
        role = st.selectbox("role", ["client", "admin"])
    with c2:
        client_id = st.text_input("client_id", placeholder="CLIENT3")
        active = st.selectbox("active", ["TRUE", "FALSE"])
        new_password = st.text_input("set / reset password", type="password", placeholder="Leave empty to keep existing")

    if st.button("Save user", type="primary"):
        if not username.strip():
            st.error("username is required.")
            return

        username_v = username.strip()
        ws = ws_users()

        # Build updated df
        df = users_df.copy()
        exists = df["username"].astype(str).eq(username_v).any()

        # Password hash logic
        if new_password.strip():
            pw_hash = bcrypt.hashpw(new_password.encode(), bcrypt.gensalt()).decode()
        else:
            if exists:
                pw_hash = df.loc[df["username"] == username_v, "password_hash"].iloc[0]
            else:
                st.error("New user requires a password.")
                return

        row_data = {
            "username": username_v,
            "name": name.strip(),
            "email": email.strip(),
            "password_hash": pw_hash,
            "role": role,
            "client_id": client_id.strip() or "DEFAULT",
            "active": active,
        }

        if exists:
            df.loc[df["username"] == username_v, :] = pd.Series(row_data)
        else:
            df = pd.concat([df, pd.DataFrame([row_data])], ignore_index=True)

        # Write back to sheet
        out = [list(df.columns)] + df.astype(str).values.tolist()
        ws.clear()
        ws.update(out)

        audit_log("admin_user_saved", {"username": username_v, "role": role, "client_id": row_data["client_id"]})
        st.success("User saved. (They can login immediately)")
        st.rerun()

    st.markdown("</div>", unsafe_allow_html=True)

    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown("### Audit Logs")
    st.markdown('<div class="muted">Stored in Google Sheet tab: <b>Audit</b></div>', unsafe_allow_html=True)
    try:
        wsa = ws_audit()
        vals = wsa.get_all_values()
        if len(vals) >= 2:
            df_a = pd.DataFrame(vals[1:], columns=vals[0])
            st.dataframe(df_a.tail(100), use_container_width=True, hide_index=True)
        else:
            st.info("No audit records yet.")
    except Exception as e:
        st.error(f"Could not load Audit sheet: {e}")
    st.markdown("</div>", unsafe_allow_html=True)


# -----------------------------------
# MESSAGES placeholder
# -----------------------------------
def messages_page():
    st.subheader("Messages Dashboard")
    st.info("Messages dashboard coming soon")
    audit_log("view_messages", {})


# -----------------------------------
# DASHBOARD (Leads)
# -----------------------------------
def load_manual_from_gsheet() -> pd.DataFrame:
    ws = ws_data()
    values = ws.get_all_values()
    if not values or len(values) < 2:
        return pd.DataFrame(columns=MANUAL_UI_COLS)
    header = values[0]
    rows = values[1:]
    df = pd.DataFrame(rows, columns=header)
    for c in MANUAL_UI_COLS:
        if c not in df.columns:
            df[c] = ""
    df = df[MANUAL_UI_COLS].copy()
    for c in ["Impressions", "Spent (GBP)", "Leads", "Converted Leads"]:
        df[c] = pd.to_numeric(df[c], errors="coerce")
    df["Month"] = df["Month"].astype(str).str.strip()
    df = df.replace({np.nan: ""})
    return df

def save_manual_to_gsheet(df_ui: pd.DataFrame):
    ws = ws_data()
    df_ui = df_ui.copy()
    for c in MANUAL_UI_COLS:
        if c not in df_ui.columns:
            df_ui[c] = ""
    df_ui = df_ui[MANUAL_UI_COLS].fillna("")
    out = [MANUAL_UI_COLS] + df_ui.astype(str).values.tolist()
    ws.clear()
    ws.update(out)

def dashboard_page():
    st.subheader("Leads Dashboard")

    # Mode switch (still keep your original behavior)
    st.caption("Choose data source, apply filters, edit conversion values, export reports.")

    data_source = st.radio(
        "Choose data source",
        ["Upload Excel/CSV", "Enter data manually (Google Sheet)"],
        horizontal=True,
    )

    df_base = None

    if data_source == "Upload Excel/CSV":
        uploaded_file = st.file_uploader("Upload Excel / CSV", type=["xlsx", "csv"])
        if not uploaded_file:
            st.info("Upload an Excel/CSV file to continue.")
            return

        try:
            if uploaded_file.name.lower().endswith(".csv"):
                raw = pd.read_csv(uploaded_file)
            else:
                raw = pd.read_excel(uploaded_file)
            df_base = normalize_columns(raw)
        except Exception as e:
            st.error(f"Could not read your file. Reason: {e}")
            return

    else:
        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.markdown("### Manual Entry (Saved to Google Sheet)")
        st.markdown(
            '<div class="muted">Admin or internal team can enter monthly rows here and click <b>Save</b>.</div>',
            unsafe_allow_html=True
        )

        a1, a2, a3 = st.columns([1, 1, 2])
        load_btn = a1.button("Load from Google Sheet")
        save_btn = a2.button("Save to Google Sheet", type="primary")

        if load_btn:
            try:
                st.session_state.manual_data = load_manual_from_gsheet()
                st.success("Loaded from Google Sheet.")
                audit_log("manual_load", {})
                st.rerun()
            except Exception as e:
                st.error(f"Google Sheet load failed: {e}")

        manual_editor = st.data_editor(
            st.session_state.manual_data,
            use_container_width=True,
            num_rows="dynamic",
            hide_index=True,
            column_config={
                "Month": st.column_config.SelectboxColumn("Month", options=MONTH_ORDER, required=True),
                "Brand": st.column_config.TextColumn("Brand", required=True),
                "Destination": st.column_config.TextColumn("Destination", required=True),
                "Impressions": st.column_config.NumberColumn("Impressions", min_value=0, step=1),
                "Spent (GBP)": st.column_config.NumberColumn("Spent (GBP)", min_value=0.0, step=0.01),
                "Leads": st.column_config.NumberColumn("Leads", min_value=0, step=1),
                "Converted Leads": st.column_config.NumberColumn("Converted Leads", min_value=0, step=1),
            },
            key="manual_editor",
        )
        st.session_state.manual_data = manual_editor.copy()

        b1, b2, b3 = st.columns([1, 1, 2])
        if b1.button("Clear manual data"):
            st.session_state.manual_data = pd.DataFrame(columns=MANUAL_UI_COLS)
            st.success("Manual data cleared.")
            audit_log("manual_clear", {})
            st.rerun()

        b2.download_button(
            "Download manual data (CSV)",
            data=st.session_state.manual_data.to_csv(index=False).encode("utf-8"),
            file_name="manual_data.csv",
            mime="text/csv",
        )

        if save_btn:
            try:
                save_manual_to_gsheet(st.session_state.manual_data)
                st.success("Saved to Google Sheet (permanent).")
                audit_log("manual_save", {"rows": int(len(st.session_state.manual_data))})
            except Exception as e:
                st.error(f"Google Sheet save failed: {e}")

        st.markdown("</div>", unsafe_allow_html=True)

        ui = st.session_state.manual_data.copy()
        if ui.empty:
            st.info("Add at least 1 row to continue.")
            return

        try:
            df_base = normalize_columns(ui.rename(columns={
                "Month": "Month",
                "Brand": "Brand",
                "Destination": "Destination",
                "Impressions": "Impressions",
                "Spent (GBP)": "Spent (GBP)",
                "Leads": "Leads",
                "Converted Leads": "Converted Leads",
            }))
        except Exception as e:
            st.error(f"Manual data error: {e}")
            return

    # Multi-tenant filter (SaaS)
    if user["role"] != "admin":
        df_base = df_base[df_base["client_id"].astype(str) == str(user["client_id"])]

    df_base = compute_metrics(df_base)
    if df_base.empty:
        st.warning("No data available for your client_id.")
        return

    # Filters
    st.sidebar.header("Filters")
    available_months = [m for m in MONTH_ORDER if m in df_base["month"].dropna().astype(str).unique().tolist()]
    if not available_months:
        available_months = sorted(df_base["month"].dropna().astype(str).unique().tolist())

    month = st.sidebar.selectbox("Month", available_months)
    d0 = df_base[df_base["month"].astype(str) == str(month)].copy()

    brand = st.sidebar.selectbox("Brand", ["All"] + sorted(d0["brand"].dropna().unique()))
    if brand != "All":
        d0 = d0[d0["brand"] == brand]

    destination = st.sidebar.selectbox("Destination", ["All"] + sorted(d0["destination"].dropna().unique()))
    if destination != "All":
        d0 = d0[d0["destination"] == destination]

    d0["row_id"] = d0["month"].astype(str) + "||" + d0["brand"].astype(str) + "||" + d0["destination"].astype(str)
    d0 = apply_overrides(d0, st.session_state.overrides)
    d = compute_metrics(d0)

    # Export CSV
    csv_data = d[FINAL_COL_ORDER].to_csv(index=False).encode("utf-8")
    if st.download_button(
        "Download filtered data (CSV)",
        data=csv_data,
        file_name=f"filtered_{month}_{brand}_{destination}.csv".replace(" ", "_"),
        mime="text/csv",
    ):
        audit_log("download_csv", {"month": month, "brand": brand, "destination": destination})

    st.write("")

    # KPI row (Spend/Leads/Brands/Dests) before table
    total_spend = float(d["spent_gbp"].sum())
    total_leads = int(d["leads"].sum())
    kpi1, kpi2, kpi3, kpi4 = st.columns(4)
    kpi1.metric("Total Spend", f"£{total_spend:,.2f}")
    kpi2.metric("Total Leads", f"{total_leads:,}")
    kpi3.metric("Brands", f"{d['brand'].nunique():,}")
    kpi4.metric("Destinations", f"{d['destination'].nunique():,}")
    st.divider()

    # Top Brands before table
    st.markdown("## Top Brands")
    left, right = st.columns(2)

    top_spend = (
        d.groupby("brand", as_index=False)["spent_gbp"].sum()
        .sort_values("spent_gbp", ascending=False)
        .head(3)
    )
    top_leads = (
        d.groupby("brand", as_index=False)["leads"].sum()
        .sort_values("leads", ascending=False)
        .head(3)
    )

    with left:
        st.markdown("### Top 3 by Spend")
        for i, row in enumerate(top_spend.itertuples(index=False), start=1):
            st.markdown(f"**#{i} {row.brand} — £{row.spent_gbp:,.2f}**")

    with right:
        st.markdown("### Top 3 by Leads")
        for i, row in enumerate(top_leads.itertuples(index=False), start=1):
            st.markdown(f"**#{i} {row.brand} — {int(row.leads):,} leads**")

    st.divider()

    # Editable Table
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown("### Editable Table (Impressions + Converted Leads)")
    st.markdown(
        '<div class="muted">Edit <b>Impressions</b> and <b>Converted Leads</b>, then click <b>Apply changes</b>.</div>',
        unsafe_allow_html=True
    )

    template_df = pd.DataFrame([
        {"Month":"January","Brand":"OWT","Destination":"Philippines","Impressions":0,"Spent (GBP)":0,"Leads":0,"Converted Leads":0},
        {"Month":"January","Brand":"OWT","Destination":"Thailand","Impressions":0,"Spent (GBP)":0,"Leads":0,"Converted Leads":0},
    ])
    st.download_button(
        "Download CSV Template",
        data=template_df.to_csv(index=False).encode("utf-8"),
        file_name="leads_template.csv",
        mime="text/csv",
    )

    with st.form("edit_form", clear_on_submit=False):
        table = d[FINAL_COL_ORDER].copy()
        table["conversion_rate"] = (table["conversion_rate"] * 100).round(2)

        edited = st.data_editor(
            table,
            use_container_width=True,
            hide_index=True,
            column_config={
                "month": st.column_config.TextColumn("Month", disabled=True),
                "brand": st.column_config.TextColumn("Brand", disabled=True),
                "destination": st.column_config.TextColumn("Destination", disabled=True),
                "impressions": st.column_config.NumberColumn("Impressions", min_value=0, step=1),
                "cpl": st.column_config.NumberColumn("CPL", format="£%.2f", disabled=True),
                "spent_gbp": st.column_config.NumberColumn("Spent (GBP)", format="£%.2f", disabled=True),
                "leads": st.column_config.NumberColumn("Leads", disabled=True),
                "converted_leads": st.column_config.NumberColumn("Converted Leads", min_value=0, step=1),
                "conversion_rate": st.column_config.NumberColumn("Conversion Rate", format="%.2f%%", disabled=True),
            },
        )

        cA, cB = st.columns([1, 1])
        apply_btn = cA.form_submit_button("Apply changes", type="primary")
        reset_btn = cB.form_submit_button("Reset manual edits")

    if reset_btn:
        st.session_state.overrides = {"impressions": {}, "converted_leads": {}}
        audit_log("reset_edits", {"month": month, "brand": brand, "destination": destination})
        st.success("Manual edits cleared.")
        st.rerun()

    if apply_btn:
        edited_row_id = (
            edited["month"].astype(str)
            + "||" + edited["brand"].astype(str)
            + "||" + edited["destination"].astype(str)
        )

        imp_map = st.session_state.overrides["impressions"]
        conv_map = st.session_state.overrides["converted_leads"]

        for rid, imp, conv in zip(edited_row_id, edited["impressions"], edited["converted_leads"]):
            imp_map[str(rid)] = int(pd.to_numeric(imp, errors="coerce") or 0)
            conv_map[str(rid)] = int(pd.to_numeric(conv, errors="coerce") or 0)

        st.session_state.overrides["impressions"] = imp_map
        st.session_state.overrides["converted_leads"] = conv_map

        audit_log("apply_changes", {"month": month, "brand": brand, "destination": destination})
        st.success("Changes applied.")
        st.rerun()

    st.markdown("</div>", unsafe_allow_html=True)
    st.divider()

    # 6 KPI summary BELOW the table (your requested placement)
    total_impr = int(d["impressions"].sum())
    total_conv = int(d["converted_leads"].sum())
    overall_cpl = (total_spend / total_leads) if total_leads > 0 else 0.0
    overall_cr = (total_conv / total_leads) if total_leads > 0 else 0.0

    k1, k2, k3, k4, k5, k6 = st.columns(6)
    k1.metric("Total Spend", f"£{total_spend:,.2f}")
    k2.metric("Total Leads", f"{total_leads:,}")
    k3.metric("Total Impressions", f"{total_impr:,}")
    k4.metric("Overall CPL", f"£{overall_cpl:,.2f}")
    k5.metric("Total Converted Leads", f"{total_conv:,}")
    k6.metric("Overall Conversion Rate", f"{overall_cr*100:,.2f}%")

    st.divider()

    # Charts
    st.subheader("Brand Level")
    b1, b2, b3 = st.columns(3)

    brand_spend = d.groupby("brand", as_index=False)["spent_gbp"].sum().sort_values("spent_gbp", ascending=True)
    fig1 = px.bar(brand_spend, x="spent_gbp", y="brand", orientation="h", title="Spend by Brand")
    b1.plotly_chart(fig1, use_container_width=True)

    brand_leads = d.groupby("brand", as_index=False)["leads"].sum().sort_values("leads", ascending=True)
    fig2 = px.bar(brand_leads, x="leads", y="brand", orientation="h", title="Leads by Brand")
    b2.plotly_chart(fig2, use_container_width=True)

    brand_cpl = d.groupby("brand", as_index=False).agg(spend=("spent_gbp", "sum"), leads=("leads", "sum"))
    brand_cpl["cpl"] = np.where(brand_cpl["leads"] > 0, brand_cpl["spend"] / brand_cpl["leads"], 0.0)
    brand_cpl = brand_cpl.sort_values("cpl", ascending=True)
    fig3 = px.bar(brand_cpl, x="cpl", y="brand", orientation="h", title="CPL by Brand")
    b3.plotly_chart(fig3, use_container_width=True)

    st.divider()

    st.subheader("Destination Level")
    top_n = st.slider("Top N Destinations", 5, 30, 10)

    dest = d.groupby("destination", as_index=False).agg(
        spend=("spent_gbp", "sum"),
        leads=("leads", "sum"),
        converted=("converted_leads", "sum"),
        impressions=("impressions", "sum"),
    )
    dest["conversion_rate"] = np.where(dest["leads"] > 0, dest["converted"] / dest["leads"], 0.0)

    c1, c2 = st.columns(2)
    top_spend_d = dest.sort_values("spend", ascending=False).head(top_n).sort_values("spend")
    fig4 = px.bar(top_spend_d, x="spend", y="destination", orientation="h", title="Top Destinations by Spend")
    c1.plotly_chart(fig4, use_container_width=True)

    top_leads_d = dest.sort_values("leads", ascending=False).head(top_n).sort_values("leads")
    fig5 = px.bar(top_leads_d, x="leads", y="destination", orientation="h", title="Top Destinations by Leads")
    c2.plotly_chart(fig5, use_container_width=True)

    st.divider()

    # PDF
    if PDF_AVAILABLE:
        filters = {"month": month, "brand": brand, "destination": destination}
        pdf_bytes = build_pdf_report(filters, d, user)
        pdf_name = f"report_{month}_{brand}_{destination}.pdf".replace(" ", "_")
        if st.download_button(
            "Download final summary as PDF",
            data=pdf_bytes,
            file_name=pdf_name,
            mime="application/pdf",
        ):
            audit_log("download_pdf", {"month": month, "brand": brand, "destination": destination})
    else:
        st.warning("PDF export disabled. Add 'reportlab' to requirements.txt.")


# -----------------------------------
# Router
# -----------------------------------
if menu == "Dashboard":
    dashboard_page()
elif menu == "Messages":
    messages_page()
else:
    settings_page()
