# app.py ✅ Plotly / Production version + Google Sheets Save/Load (Permanent)
# ------------------------------------------------------------
# What you get:
# 1) LEADS + MESSAGES buttons (top-center)
# 2) Data Input:
#    - Upload Excel/CSV
#    - Enter data manually
# 3) Manual data:
#    - Add/edit rows
#    - Save to Google Sheet (permanent)
#    - Load from Google Sheet
#    - Download CSV backup
# 4) Leads dashboard:
#    - Sidebar filters (Month/Brand/Destination)
#    - KPI Summary
#    - Charts (Brand + Destination)
#    - Decomposition view (tabs)
#    - Export CSV + PDF
# ------------------------------------------------------------

import io
import numpy as np
import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.io as pio

# Optional PDF
PDF_AVAILABLE = True
try:
    from reportlab.lib.pagesizes import A4
    from reportlab.pdfgen import canvas
    from reportlab.lib.units import cm
except ModuleNotFoundError:
    PDF_AVAILABLE = False

# Optional Google Sheets
GSHEETS_AVAILABLE = True
try:
    import gspread
    from google.oauth2.service_account import Credentials
except ModuleNotFoundError:
    GSHEETS_AVAILABLE = False


# -----------------------------
# Page setup + theme
# -----------------------------
st.set_page_config(page_title="Marketing Performance Dashboard", layout="wide")
pio.templates.default = "plotly_dark"

st.markdown(
    """
    <style>
      .block-container { padding-top: 1.2rem; padding-bottom: 2rem; }
      section[data-testid="stSidebar"] { padding-top: 1rem; }
      h1, h2, h3 { letter-spacing: 0.2px; }
      [data-testid="stCaptionContainer"] { opacity: 0.85; }
      [data-testid="stMetricValue"] { font-size: 1.6rem; }
      .mode-note { text-align:center; opacity:0.85; margin: 6px 0 14px 0; }
      .card { border:1px solid rgba(255,255,255,0.10); border-radius:14px; padding:14px; background:rgba(255,255,255,0.02); }
      .muted { opacity:0.85; }
    </style>
    """,
    unsafe_allow_html=True,
)

MONTH_ORDER = [
    "January","February","March","April","May","June",
    "July","August","September","October","November","December"
]

# Internal final order (matches your requirement)
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

# Manual entry (user-facing) columns (keeps it simple)
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
    """Normalize to internal column names + ensure required columns exist."""
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

    # Case-insensitive mapping
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

    # Defaults
    if "impressions" not in df.columns:
        df["impressions"] = 0
    if "converted_leads" not in df.columns:
        df["converted_leads"] = 0

    # Clean text
    df["month"] = df["month"].astype(str).str.strip()
    df["brand"] = df["brand"].astype(str).str.strip()
    df["destination"] = df["destination"].astype(str).str.strip()

    # Clean numerics
    df["impressions"] = pd.to_numeric(df["impressions"], errors="coerce").fillna(0).astype(int)
    df["spent_gbp"] = pd.to_numeric(df["spent_gbp"], errors="coerce").fillna(0.0).astype(float)
    df["leads"] = pd.to_numeric(df["leads"], errors="coerce").fillna(0).astype(int)
    df["converted_leads"] = pd.to_numeric(df["converted_leads"], errors="coerce").fillna(0).astype(int)

    # Month ordering
    if df["month"].isin(MONTH_ORDER).any():
        df["month"] = pd.Categorical(df["month"], categories=MONTH_ORDER, ordered=True)

    return df


def compute_metrics(df: pd.DataFrame) -> pd.DataFrame:
    """Required formulas:
       CPL = Spent / Leads
       Conversion Rate = Converted Leads / Leads
    """
    out = df.copy()
    leads_safe = out["leads"].replace(0, np.nan)

    out["cpl"] = (out["spent_gbp"] / leads_safe).fillna(0.0)
    out["conversion_rate"] = (out["converted_leads"] / leads_safe).fillna(0.0)
    return out


def build_pdf_report(filters: dict, d: pd.DataFrame) -> bytes:
    """Simple PDF summary for filtered data."""
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    w, h = A4

    c.setFont("Helvetica-Bold", 16)
    c.drawString(2 * cm, h - 2 * cm, "Leads Performance Report")

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
    """
    Requires these secrets in Streamlit Cloud:
    - st.secrets["gcp_service_account"]  (full service account JSON as a dict)
    - st.secrets["gsheet_id"] (spreadsheet ID)
    - optional: st.secrets["gsheet_tab"] (tab name, default "Data")
    """
    if not GSHEETS_AVAILABLE:
        raise RuntimeError("Google Sheets libraries missing. Add gspread + google-auth to requirements.txt")

    if "gcp_service_account" not in st.secrets or "gsheet_id" not in st.secrets:
        raise RuntimeError(
            "Missing Streamlit secrets. Add 'gcp_service_account' and 'gsheet_id' in Streamlit secrets."
        )

    creds_info = dict(st.secrets["gcp_service_account"])
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = Credentials.from_service_account_info(creds_info, scopes=scopes)
    client = gspread.authorize(creds)
    return client


def load_manual_from_gsheet() -> pd.DataFrame:
    client = _get_gs_client()
    sheet_id = st.secrets["gsheet_id"]
    tab = st.secrets.get("gsheet_tab", "Data")

    sh = client.open_by_key(sheet_id)
    ws = sh.worksheet(tab)

    values = ws.get_all_values()
    if not values or len(values) < 2:
        return pd.DataFrame(columns=MANUAL_UI_COLS)

    header = values[0]
    rows = values[1:]

    df = pd.DataFrame(rows, columns=header)

    # Ensure expected UI columns exist
    for c in MANUAL_UI_COLS:
        if c not in df.columns:
            df[c] = ""

    df = df[MANUAL_UI_COLS].copy()

    # Clean numeric columns
    for c in ["Impressions", "Spent (GBP)", "Leads", "Converted Leads"]:
        df[c] = pd.to_numeric(df[c], errors="coerce")

    # Month clean
    df["Month"] = df["Month"].astype(str).str.strip()

    # Replace NaNs with blanks for editor display
    df = df.replace({np.nan: ""})
    return df


def save_manual_to_gsheet(df_ui: pd.DataFrame):
    client = _get_gs_client()
    sheet_id = st.secrets["gsheet_id"]
    tab = st.secrets.get("gsheet_tab", "Data")

    sh = client.open_by_key(sheet_id)
    ws = sh.worksheet(tab)

    df_ui = df_ui.copy()

    # Keep only required UI columns
    for c in MANUAL_UI_COLS:
        if c not in df_ui.columns:
            df_ui[c] = ""
    df_ui = df_ui[MANUAL_UI_COLS]

    # Convert to strings for Sheets
    out = [MANUAL_UI_COLS] + df_ui.fillna("").astype(str).values.tolist()

    ws.clear()
    ws.update(out)


# -----------------------------
# Session state
# -----------------------------
if "mode" not in st.session_state:
    st.session_state.mode = None

if "manual_data" not in st.session_state:
    st.session_state.manual_data = pd.DataFrame(columns=MANUAL_UI_COLS)


# -----------------------------
# Header + Mode Selector
# -----------------------------
st.title("Marketing Performance Dashboard")
st.caption("Private performance dashboard — upload Excel/CSV and explore insights.")

left, mid, right = st.columns([1, 3, 1])
with mid:
    b1, b2 = st.columns(2)
    if b1.button("LEADS", use_container_width=True):
        st.session_state.mode = "LEADS"
    if b2.button("MESSAGES", use_container_width=True):
        st.session_state.mode = "MESSAGES"

st.markdown('<div class="mode-note">Click <b>LEADS</b> or <b>MESSAGES</b> to continue.</div>', unsafe_allow_html=True)

if not st.session_state.mode:
    st.stop()

st.divider()

# -----------------------------
# Upload / Manual input
# -----------------------------
st.subheader("Data Input")

data_source = st.radio(
    "Choose data source",
    ["Upload Excel/CSV", "Enter data manually"],
    horizontal=True
)

df = None  # will be prepared for the dashboard

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
        df = normalize_columns(raw)
    except Exception as e:
        st.error(f"Could not read your file. Reason: {e}")
        st.stop()

else:
    # Manual entry UI
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown("### Manual Entry (Permanent Save to Google Sheets)")
    st.markdown(
        '<div class="muted">Add rows + fill Month, Brand, Destination, Spent (GBP), Leads. '
        'Converted Leads optional. Use <b>Save to Google Sheet</b> to keep it permanently.</div>',
        unsafe_allow_html=True
    )

    # Google Sheet actions
    g1, g2, g3 = st.columns([1, 1, 2])
    load_btn = g1.button("Load from Google Sheet")
    save_btn = g2.button("Save to Google Sheet", type="primary")

    if not GSHEETS_AVAILABLE:
        st.warning("Google Sheets not enabled. Add `gspread` and `google-auth` to requirements.txt.")
    else:
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
        key="manual_editor"
    )

    # Save to session (so it doesn't vanish on rerun)
    st.session_state.manual_data = manual_editor.copy()

    c1, c2, c3 = st.columns([1, 1, 2])
    clear_manual = c1.button("Clear manual data")
    download_csv = c2.download_button(
        "Download manual data (CSV)",
        data=st.session_state.manual_data.to_csv(index=False).encode("utf-8"),
        file_name="manual_data.csv",
        mime="text/csv"
    )

    if clear_manual:
        st.session_state.manual_data = pd.DataFrame(columns=MANUAL_UI_COLS)
        st.success("Manual data cleared.")
        st.rerun()

    if save_btn:
        if not GSHEETS_AVAILABLE:
            st.error("Google Sheets libraries missing. Add gspread + google-auth.")
        else:
            try:
                save_manual_to_gsheet(st.session_state.manual_data)
                st.success("Saved to Google Sheet (permanent).")
            except Exception as e:
                st.error(f"Google Sheet save failed: {e}")

    st.markdown("</div>", unsafe_allow_html=True)

    # Convert manual schema -> expected schema for dashboard
    ui = st.session_state.manual_data.copy()

    # basic validation
    if ui.empty:
        st.info("Add at least 1 row to continue.")
        st.stop()

    # Rename to internal then normalize
    ui2 = ui.rename(columns={
        "Month": "Month",
        "Brand": "Brand",
        "Destination": "Destination",
        "Impressions": "Impressions",
        "Spent (GBP)": "Spent (GBP)",
        "Leads": "Leads",
        "Converted Leads": "Converted Leads",
    })

    try:
        df = normalize_columns(ui2)
    except Exception as e:
        st.error(f"Manual data error: {e}")
        st.stop()

st.divider()

# -----------------------------
# Messages dashboard placeholder
# -----------------------------
if st.session_state.mode == "MESSAGES":
    st.subheader("Messages Dashboard")
    st.info("Messages dashboard coming soon")
    st.stop()

# -----------------------------
# Leads dashboard filters
# -----------------------------
df = compute_metrics(df)

st.sidebar.header("Filters")

available_months = [m for m in MONTH_ORDER if m in df["month"].dropna().astype(str).unique().tolist()]
if not available_months:
    available_months = sorted(df["month"].dropna().astype(str).unique().tolist())

month = st.sidebar.selectbox("Month", available_months)
d = df[df["month"].astype(str) == str(month)].copy()

brand = st.sidebar.selectbox("Brand", ["All"] + sorted(d["brand"].dropna().unique()))
if brand != "All":
    d = d[d["brand"] == brand]

destination = st.sidebar.selectbox("Destination", ["All"] + sorted(d["destination"].dropna().unique()))
if destination != "All":
    d = d[d["destination"] == destination]

d = compute_metrics(d)

# -----------------------------
# Export filtered CSV
# -----------------------------
st.download_button(
    "Download filtered data (CSV)",
    d[FINAL_COL_ORDER].to_csv(index=False).encode("utf-8"),
    file_name=f"filtered_{month}_{brand}_{destination}.csv".replace(" ", "_"),
    mime="text/csv",
)

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
# Charts: Brand Level
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
# Charts: Destination Level
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
# Decomposition View
# -----------------------------
st.subheader("Decomposition View")

t1, t2, t3 = st.tabs(["Brand summary", "Destination summary", "Brand → Destination detail"])

with t1:
    bs = d.groupby("brand", as_index=False).agg(
        spend=("spent_gbp", "sum"),
        leads=("leads", "sum"),
        impressions=("impressions", "sum"),
        converted=("converted_leads", "sum"),
    )
    bs["cpl"] = np.where(bs["leads"] > 0, bs["spend"] / bs["leads"], 0.0)
    bs["conversion_rate_%"] = np.where(bs["leads"] > 0, (bs["converted"] / bs["leads"]) * 100, 0.0).round(2)
    st.dataframe(bs.sort_values("spend", ascending=False), use_container_width=True, hide_index=True)

with t2:
    ds = dest.copy()
    ds["conversion_rate_%"] = (ds["conversion_rate"] * 100).round(2)
    ds = ds.drop(columns=["conversion_rate"])
    st.dataframe(ds.sort_values("spend", ascending=False), use_container_width=True, hide_index=True)

with t3:
    bd = d.groupby(["brand", "destination"], as_index=False).agg(
        spend=("spent_gbp", "sum"),
        leads=("leads", "sum"),
        impressions=("impressions", "sum"),
        converted=("converted_leads", "sum"),
    )
    bd["cpl"] = np.where(bd["leads"] > 0, bd["spend"] / bd["leads"], 0.0)
    bd["conversion_rate_%"] = np.where(bd["leads"] > 0, (bd["converted"] / bd["leads"]) * 100, 0.0).round(2)
    st.dataframe(bd.sort_values(["spend", "leads"], ascending=False), use_container_width=True, hide_index=True)

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
# Optional: final dataset
# -----------------------------
with st.expander("Show final filtered dataset"):
    out = d[FINAL_COL_ORDER].copy()
    out["conversion_rate_%"] = (out["conversion_rate"] * 100).round(2)
    st.dataframe(out, use_container_width=True, hide_index=True)
