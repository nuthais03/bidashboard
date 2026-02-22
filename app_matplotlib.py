# app.py ✅ FULL PRODUCTION VERSION (OWT Branded + Light/Dark Toggle + Upload + Manual Entry to Google Sheets + Table + CSV Template + Charts + PDF)
# ------------------------------------------------------------------------------------------------------------------
# requirements.txt:
# streamlit
# pandas
# numpy
# plotly
# openpyxl
# reportlab
# gspread
# google-auth
#
# Streamlit Secrets (Settings -> Secrets):
# gsheet_id = "YOUR_SPREADSHEET_ID"
# gsheet_tab = "Data"   # optional (default: Data)
#
# [gcp_service_account]
# type = "service_account"
# project_id = "..."
# private_key_id = "..."
# private_key = "-----BEGIN PRIVATE KEY-----\n...\n-----END PRIVATE KEY-----\n"
# client_email = "....iam.gserviceaccount.com"
# client_id = "..."
# auth_uri = "https://accounts.google.com/o/oauth2/auth"
# token_uri = "https://oauth2.googleapis.com/token"
# auth_provider_x509_cert_url = "https://www.googleapis.com/oauth2/v1/certs"
# client_x509_cert_url = "..."
# ------------------------------------------------------------------------------------------------------------------

import io
import numpy as np
import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.io as pio

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


# -----------------------------
# Branding config (OWT)
# -----------------------------
OWT_LOGO_URL = "https://owtgroupltd.co.uk/assets/images/OWT%20Group%20Logo-TransparentBackground1.png"
APP_TITLE = "Marketing Performance Dashboard"
APP_TAGLINE = "Private performance dashboard — upload Excel/CSV or enter data manually."
BRAND_ACCENT = "#3B82F6"


# -----------------------------
# Page setup
# -----------------------------
st.set_page_config(page_title=APP_TITLE, layout="wide")


# -----------------------------
# Theme State
# -----------------------------
if "theme_mode" not in st.session_state:
    st.session_state.theme_mode = "Dark"


def apply_theme(theme_mode: str) -> None:
    """Apply CSS + Plotly template based on theme."""
    if theme_mode == "Light":
        pio.templates.default = "plotly_white"
        st.markdown(
            f"""
            <style>
              .block-container {{ padding-top: 1.0rem; padding-bottom: 2rem; }}
              section[data-testid="stSidebar"] {{ background-color: #ffffff; }}
              .stApp {{ background-color: #f5f7fb; color: #111111; }}

              h1,h2,h3 {{ letter-spacing: 0.2px; color:#111111; }}
              [data-testid="stCaptionContainer"] {{ opacity: 0.80; color:#111111; }}
              [data-testid="stMetricValue"] {{ font-size: 1.6rem; color:#111111; }}
              [data-testid="stMetricLabel"] {{ color:#111111; opacity:0.85; }}

              .card {{
                border:1px solid #e5e7eb;
                border-radius:14px;
                padding:14px;
                background:#ffffff;
              }}
              .muted {{ opacity:0.75; color:#111111; }}

              /* Buttons */
              div.stButton > button {{
                border-radius: 10px;
              }}

              /* Sidebar branding block separator */
              .sb-sep {{
                height:1px;
                background: rgba(17,17,17,0.08);
                margin: 0 0 10px 0;
              }}

              /* Header bar (visual) */
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
                height: 8px;
                width: 8px;
                border-radius: 50%;
                background: {BRAND_ACCENT};
                display: inline-block;
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
              [data-testid="stCaptionContainer"] {{ opacity: 0.85; }}
              [data-testid="stMetricValue"] {{ font-size: 1.6rem; }}
              .card {{
                border:1px solid rgba(255,255,255,0.10);
                border-radius:14px;
                padding:14px;
                background:rgba(255,255,255,0.02);
              }}
              .muted {{ opacity:0.85; }}

              div.stButton > button {{
                border-radius: 10px;
              }}

              .sb-sep {{
                height:1px;
                background: rgba(255,255,255,0.08);
                margin: 0 0 10px 0;
              }}

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
                height: 8px;
                width: 8px;
                border-radius: 50%;
                background: {BRAND_ACCENT};
                display: inline-block;
              }}
            </style>
            """,
            unsafe_allow_html=True,
        )


apply_theme(st.session_state.theme_mode)


# -----------------------------
# Sidebar brand block
# -----------------------------
st.sidebar.markdown(
    f"""
    <div style="display:flex; align-items:center; gap:10px; padding:8px 6px 12px 6px;">
      <img src="{OWT_LOGO_URL}" style="height:38px; width:auto;" />
      <div>
        <div style="font-weight:800; line-height:1.1;">OWT Dashboard</div>
        <div style="opacity:0.75; font-size:12px;">Leads & Messaging</div>
      </div>
    </div>
    <div class="sb-sep"></div>
    """,
    unsafe_allow_html=True,
)


# -----------------------------
# Header row (Enterprise look) + Theme toggle (Top-right)
# -----------------------------
st.markdown('<div class="header-bar">', unsafe_allow_html=True)
h1, h2, h3, h4 = st.columns([2, 6, 2, 2])

with h1:
    st.image(OWT_LOGO_URL, width=150)

with h2:
    st.markdown(f"### {APP_TITLE}")
    st.caption(APP_TAGLINE)

with h3:
    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown(f'<span class="pill"><span class="dot"></span>Private</span>', unsafe_allow_html=True)
    st.markdown("")

with h4:
    st.markdown("<br>", unsafe_allow_html=True)
    light = st.toggle("Light mode", value=(st.session_state.theme_mode == "Light"))
    new_mode = "Light" if light else "Dark"
    if new_mode != st.session_state.theme_mode:
        st.session_state.theme_mode = new_mode
        st.rerun()

st.markdown("</div>", unsafe_allow_html=True)
st.write("")  # spacing


# -----------------------------
# Constants
# -----------------------------
MONTH_ORDER = [
    "January","February","March","April","May","June",
    "July","August","September","October","November","December"
]

FINAL_COL_ORDER = [
    "month",
    "brand",
    "destination",
    "impressions",
    "cpl",
    "spent_gbp",
    "leads",
    "converted_leads",
    "conversion_rate",
]

MANUAL_UI_COLS = [
    "Month",
    "Brand",
    "Destination",
    "Impressions",
    "Spent (GBP)",
    "Leads",
    "Converted Leads",
]


# -----------------------------
# Helpers
# -----------------------------
def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = df.columns.str.strip()

    rename_map = {
        "Month": "month",
        "Brand": "brand",
        "Destination": "destination",
        "Impressions": "impressions",
        "CPL": "cpl",
        "Spent (GBP)": "spent_gbp",
        "Spent": "spent_gbp",
        "Spend": "spent_gbp",
        "Leads": "leads",
        "Converted Leads": "converted_leads",
        "Conversion Rate": "conversion_rate",
    }

    lower_to_actual = {c.lower(): c for c in df.columns}
    for k, v in list(rename_map.items()):
        if k not in df.columns and k.lower() in lower_to_actual:
            rename_map[lower_to_actual[k.lower()]] = v

    df = df.rename(columns={k: v for k, v in rename_map.items() if k in df.columns})

    required = {"month", "brand", "destination", "leads", "spent_gbp"}
    missing = required - set(df.columns)
    if missing:
        raise ValueError(
            "Missing required columns: "
            + ", ".join(sorted(missing))
            + " | Required: Month, Brand, Destination, Leads, Spent (GBP)"
        )

    if "impressions" not in df.columns:
        df["impressions"] = 0
    if "converted_leads" not in df.columns:
        df["converted_leads"] = 0

    df["month"] = df["month"].astype(str).str.strip()
    df["brand"] = df["brand"].astype(str).str.strip()
    df["destination"] = df["destination"].astype(str).str.strip()

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
        df["row_id"] = (
            df["month"].astype(str) + "||" + df["brand"].astype(str) + "||" + df["destination"].astype(str)
        )

    imp_map = overrides.get("impressions", {})
    conv_map = overrides.get("converted_leads", {})

    if imp_map:
        df["impressions"] = df["row_id"].map(imp_map).combine_first(df["impressions"]).astype(int)
    if conv_map:
        df["converted_leads"] = df["row_id"].map(conv_map).combine_first(df["converted_leads"]).astype(int)

    return df


def build_pdf_report(filters: dict, d: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    w, h = A4

    c.setFont("Helvetica-Bold", 16)
    c.drawString(2 * cm, h - 2 * cm, "OWT — Leads Performance Report")

    c.setFont("Helvetica", 10)
    c.drawString(
        2 * cm,
        h - 2.8 * cm,
        f"Month: {filters.get('month','All')} | Brand: {filters.get('brand','All')} | Destination: {filters.get('destination','All')}",
    )

    total_spend = float(d["spent_gbp"].sum())
    total_leads = int(d["leads"].sum())
    total_impr = int(d["impressions"].sum())
    total_conv = int(d["converted_leads"].sum())
    overall_cpl = (total_spend / total_leads) if total_leads > 0 else 0.0
    overall_cr = (total_conv / total_leads) if total_leads > 0 else 0.0

    y = h - 4.0 * cm
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

    y -= 1.2 * cm
    c.setFont("Helvetica-Bold", 12)
    c.drawString(2 * cm, y, "Top Rows (by Leads)")
    y -= 0.7 * cm

    top = d.sort_values("leads", ascending=False).head(12).copy()
    top["cr_pct"] = top["conversion_rate"] * 100

    c.setFont("Helvetica-Bold", 9)
    c.drawString(2 * cm, y, "Brand")
    c.drawString(6.2 * cm, y, "Destination")
    c.drawRightString(12.6 * cm, y, "Spend")
    c.drawRightString(15.0 * cm, y, "Leads")
    c.drawRightString(17.0 * cm, y, "CPL")
    c.drawRightString(19.5 * cm, y, "CR%")
    y -= 0.5 * cm

    c.setFont("Helvetica", 9)
    for _, r in top.iterrows():
        if y < 2.2 * cm:
            c.showPage()
            y = h - 2.0 * cm
            c.setFont("Helvetica", 9)

        c.drawString(2 * cm, y, str(r["brand"])[:18])
        c.drawString(6.2 * cm, y, str(r["destination"])[:20])
        c.drawRightString(12.6 * cm, y, f"£{float(r['spent_gbp']):,.2f}")
        c.drawRightString(15.0 * cm, y, f"{int(r['leads']):,}")
        c.drawRightString(17.0 * cm, y, f"£{float(r['cpl']):,.2f}")
        c.drawRightString(19.5 * cm, y, f"{float(r['cr_pct']):.2f}%")
        y -= 0.45 * cm

    c.showPage()
    c.save()
    buf.seek(0)
    return buf.read()


# -----------------------------
# Google Sheets helpers
# -----------------------------
def _get_gs_client():
    if not GSHEETS_AVAILABLE:
        raise RuntimeError("Google Sheets libraries missing. Add gspread + google-auth to requirements.txt")

    if "gcp_service_account" not in st.secrets or "gsheet_id" not in st.secrets:
        raise RuntimeError("Missing secrets: add gcp_service_account + gsheet_id in Streamlit secrets.")

    creds_info = dict(st.secrets["gcp_service_account"])
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = Credentials.from_service_account_info(creds_info, scopes=scopes)
    return gspread.authorize(creds)


def gsheet_get_ws():
    client = _get_gs_client()
    sheet_id = st.secrets["gsheet_id"]
    tab = st.secrets.get("gsheet_tab", "Data")
    sh = client.open_by_key(sheet_id)
    return sh.worksheet(tab)


def load_manual_from_gsheet() -> pd.DataFrame:
    ws = gsheet_get_ws()
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
    ws = gsheet_get_ws()

    df_ui = df_ui.copy()
    for c in MANUAL_UI_COLS:
        if c not in df_ui.columns:
            df_ui[c] = ""
    df_ui = df_ui[MANUAL_UI_COLS].fillna("")

    out = [MANUAL_UI_COLS] + df_ui.astype(str).values.tolist()
    ws.clear()
    ws.update(out)


# -----------------------------
# Session state
# -----------------------------
if "mode" not in st.session_state:
    st.session_state.mode = None

if "overrides" not in st.session_state:
    st.session_state.overrides = {"impressions": {}, "converted_leads": {}}

if "manual_data" not in st.session_state:
    st.session_state.manual_data = pd.DataFrame(columns=MANUAL_UI_COLS)


# -----------------------------
# Mode selector (centered)
# -----------------------------
sp_l, mid, sp_r = st.columns([1, 3, 1])
with mid:
    b1, b2 = st.columns(2)
    if b1.button("LEADS", use_container_width=True):
        st.session_state.mode = "LEADS"
    if b2.button("MESSAGES", use_container_width=True):
        st.session_state.mode = "MESSAGES"

st.markdown(
    '<div style="text-align:center; opacity:0.85; margin: 6px 0 14px 0;">'
    'Click <b>LEADS</b> or <b>MESSAGES</b> to continue.'
    "</div>",
    unsafe_allow_html=True,
)

if not st.session_state.mode:
    st.stop()

st.divider()


# -----------------------------
# Data input
# -----------------------------
st.subheader("Data Input")

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
        st.stop()

    try:
        if uploaded_file.name.lower().endswith(".csv"):
            raw = pd.read_csv(uploaded_file)
        else:
            raw = pd.read_excel(uploaded_file)
        df_base = normalize_columns(raw)
    except Exception as e:
        st.error(f"Could not read your file. Reason: {e}")
        st.stop()

else:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown("### Manual Entry (Saved to Google Sheet)")
    st.markdown(
        '<div class="muted">Add rows and fill Month, Brand, Destination, Spent (GBP), Leads. '
        'Converted Leads optional. Click <b>Save to Google Sheet</b> to store permanently.</div>',
        unsafe_allow_html=True
    )

    a1, a2, a3 = st.columns([1, 1, 2])
    load_btn = a1.button("Load from Google Sheet")
    save_btn = a2.button("Save to Google Sheet", type="primary")

    if load_btn:
        try:
            st.session_state.manual_data = load_manual_from_gsheet()
            st.success("Loaded from Google Sheet.")
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
        except Exception as e:
            st.error(f"Google Sheet save failed: {e}")

    st.markdown("</div>", unsafe_allow_html=True)

    ui = st.session_state.manual_data.copy()
    if ui.empty:
        st.info("Add at least 1 row to continue.")
        st.stop()

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
        st.stop()

st.divider()


# -----------------------------
# Messages placeholder
# -----------------------------
if st.session_state.mode == "MESSAGES":
    st.subheader("Messages Dashboard")
    st.info("Messages dashboard coming soon")
    st.stop()


# -----------------------------
# Leads filters
# -----------------------------
df_base = compute_metrics(df_base)

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


# -----------------------------
# CSV export
# -----------------------------
st.download_button(
    "Download filtered data (CSV)",
    d[FINAL_COL_ORDER].to_csv(index=False).encode("utf-8"),
    file_name=f"filtered_{month}_{brand}_{destination}.csv".replace(" ", "_"),
    mime="text/csv",
)


# -----------------------------
# Editable Table + CSV Template
# -----------------------------
st.markdown('<div class="card">', unsafe_allow_html=True)
st.markdown("### Editable Table (Impressions + Converted Leads)")
st.markdown(
    '<div class="muted">Edit <b>Impressions</b> and <b>Converted Leads</b>, then click <b>Apply changes</b>. '
    '<b>CPL</b> and <b>Conversion Rate</b> update automatically.</div>',
    unsafe_allow_html=True
)

template_df = pd.DataFrame([
    {"Month":"January","Brand":"OWT","Destination":"Philippines","Impressions":0,"Spent (GBP)":0,"Leads":0,"Converted Leads":0},
    {"Month":"January","Brand":"TH-UK","Destination":"Thailand","Impressions":0,"Spent (GBP)":0,"Leads":0,"Converted Leads":0},
])
st.download_button(
    "Download CSV Template",
    data=template_df.to_csv(index=False).encode("utf-8"),
    file_name="leads_template.csv",
    mime="text/csv"
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

    col1, col2 = st.columns([1, 1])
    apply_btn = col1.form_submit_button("Apply changes", type="primary")
    reset_btn = col2.form_submit_button("Reset manual edits")

if reset_btn:
    st.session_state.overrides = {"impressions": {}, "converted_leads": {}}
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

    st.success("Changes applied.")
    st.rerun()

st.markdown("</div>", unsafe_allow_html=True)
st.divider()


# -----------------------------
# KPI Summary
# -----------------------------
total_spend = float(d["spent_gbp"].sum())
total_leads = int(d["leads"].sum())
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


# -----------------------------
# Charts: Brand level
# -----------------------------
st.subheader("Brand Level")
b1, b2, b3 = st.columns(3)

brand_spend = d.groupby("brand", as_index=False)["spent_gbp"].sum().sort_values("spent_gbp", ascending=True)
fig1 = px.bar(brand_spend, x="spent_gbp", y="brand", orientation="h", title="Spend by Brand")
fig1.update_layout(xaxis_title="Spend (GBP)", yaxis_title="Brand")
b1.plotly_chart(fig1, use_container_width=True)

brand_leads = d.groupby("brand", as_index=False)["leads"].sum().sort_values("leads", ascending=True)
fig2 = px.bar(brand_leads, x="leads", y="brand", orientation="h", title="Leads by Brand")
fig2.update_layout(xaxis_title="Leads", yaxis_title="Brand")
b2.plotly_chart(fig2, use_container_width=True)

brand_cpl = d.groupby("brand", as_index=False).agg(spend=("spent_gbp", "sum"), leads=("leads", "sum"))
brand_cpl["cpl"] = np.where(brand_cpl["leads"] > 0, brand_cpl["spend"] / brand_cpl["leads"], 0.0)
brand_cpl = brand_cpl.sort_values("cpl", ascending=True)
fig3 = px.bar(brand_cpl, x="cpl", y="brand", orientation="h", title="CPL by Brand")
fig3.update_layout(xaxis_title="CPL (GBP)", yaxis_title="Brand")
b3.plotly_chart(fig3, use_container_width=True)
st.divider()


# -----------------------------
# Charts: Destination level
# -----------------------------
st.subheader("Destination Level")
top_n = st.slider("Top N Destinations", 5, 30, 10)

dest = d.groupby("destination", as_index=False).agg(
    spend=("spent_gbp", "sum"),
    leads=("leads", "sum"),
    converted=("converted_leads", "sum"),
    impressions=("impressions", "sum"),
)
dest["cpl"] = np.where(dest["leads"] > 0, dest["spend"] / dest["leads"], 0.0)
dest["conversion_rate"] = np.where(dest["leads"] > 0, dest["converted"] / dest["leads"], 0.0)

c1, c2 = st.columns(2)

top_spend = dest.sort_values("spend", ascending=False).head(top_n).sort_values("spend")
fig4 = px.bar(top_spend, x="spend", y="destination", orientation="h", title="Top Destinations by Spend")
fig4.update_layout(xaxis_title="Spend (GBP)", yaxis_title="Destination")
c1.plotly_chart(fig4, use_container_width=True)

top_leads = dest.sort_values("leads", ascending=False).head(top_n).sort_values("leads")
fig5 = px.bar(top_leads, x="leads", y="destination", orientation="h", title="Top Destinations by Leads")
fig5.update_layout(xaxis_title="Leads", yaxis_title="Destination")
c2.plotly_chart(fig5, use_container_width=True)
st.divider()


# -----------------------------
# PDF Export
# -----------------------------
filters = {"month": month, "brand": brand, "destination": destination}
if PDF_AVAILABLE:
    pdf_bytes = build_pdf_report(filters, d)
    st.download_button(
        "Download final summary as PDF",
        data=pdf_bytes,
        file_name=f"report_{month}_{brand}_{destination}.pdf".replace(" ", "_"),
        mime="application/pdf",
    )
else:
    st.warning("PDF export disabled. Add 'reportlab' to requirements.txt.")


# -----------------------------
# Final dataset (optional)
# -----------------------------
with st.expander("Show final filtered dataset"):
    out = d[FINAL_COL_ORDER].copy()
    out["conversion_rate_%"] = (out["conversion_rate"] * 100).round(2)
    st.dataframe(out, use_container_width=True, hide_index=True)
