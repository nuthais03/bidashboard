import io
import numpy as np
import pandas as pd
import streamlit as st
import matplotlib.pyplot as plt

PDF_AVAILABLE = True
try:
    from reportlab.lib.pagesizes import A4
    from reportlab.pdfgen import canvas
    from reportlab.lib.units import cm
except ModuleNotFoundError:
    PDF_AVAILABLE = False

st.set_page_config(page_title="Marketing Performance Dashboard (Matplotlib)", layout="wide")

st.markdown(
    """
    <style>
      .block-container { padding-top: 1.2rem; padding-bottom: 2rem; }
      section[data-testid="stSidebar"] { padding-top: 1rem; }
      [data-testid="stMetricValue"] { font-size: 1.65rem; }
      .card { border:1px solid rgba(255,255,255,0.10); border-radius:14px; padding:14px; background:rgba(255,255,255,0.02); }
      .muted { opacity:0.85; }
      .mode-note { text-align:center; opacity:0.85; margin: 6px 0 14px 0; }
    </style>
    """,
    unsafe_allow_html=True,
)

MONTH_ORDER = [
    "January","February","March","April","May","June",
    "July","August","September","October","November","December"
]

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

    lower = {c.lower(): c for c in df.columns}
    for k, v in list(rename_map.items()):
        if k not in df.columns and k.lower() in lower:
            rename_map[lower[k.lower()]] = v

    df = df.rename(columns={k: v for k, v in rename_map.items() if k in df.columns})

    required = {"month", "brand", "destination", "leads"}
    missing = required - set(df.columns)
    if missing:
        raise ValueError(f"Missing required columns: {', '.join(sorted(missing))}")

    if "impressions" not in df.columns:
        df["impressions"] = 0
    if "spent_gbp" not in df.columns:
        df["spent_gbp"] = 0.0
    if "converted_leads" not in df.columns:
        df["converted_leads"] = 0
    if "cpl" not in df.columns:
        df["cpl"] = np.nan
    if "conversion_rate" not in df.columns:
        df["conversion_rate"] = np.nan

    df["month"] = df["month"].astype(str).str.strip()
    df["brand"] = df["brand"].astype(str).str.strip()
    df["destination"] = df["destination"].astype(str).str.strip()

    df["impressions"] = pd.to_numeric(df["impressions"], errors="coerce").fillna(0).astype(int)
    df["leads"] = pd.to_numeric(df["leads"], errors="coerce").fillna(0).astype(int)
    df["converted_leads"] = pd.to_numeric(df["converted_leads"], errors="coerce").fillna(0).astype(int)
    df["spent_gbp"] = pd.to_numeric(df["spent_gbp"], errors="coerce").fillna(0.0).astype(float)

    if df["month"].isin(MONTH_ORDER).any():
        df["month"] = pd.Categorical(df["month"], categories=MONTH_ORDER, ordered=True)

    return df


def compute_metrics(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    leads_safe = out["leads"].replace(0, np.nan)
    out["cpl"] = (out["spent_gbp"] / leads_safe).fillna(0.0)
    out["conversion_rate"] = (out["converted_leads"] / leads_safe).fillna(0.0)
    return out


def build_pdf_report(filters: dict, d: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    w, h = A4

    c.setFont("Helvetica-Bold", 16)
    c.drawString(2*cm, h - 2*cm, "Leads Performance Report (Matplotlib App)")

    c.setFont("Helvetica", 10)
    c.drawString(
        2*cm, h - 2.8*cm,
        f"Month: {filters.get('month','All')} | Brand: {filters.get('brand','All')} | Destination: {filters.get('destination','All')}"
    )

    total_spend = float(d["spent_gbp"].sum())
    total_leads = int(d["leads"].sum())
    total_impr = int(d["impressions"].sum())
    total_conv = int(d["converted_leads"].sum())
    overall_cpl = (total_spend / total_leads) if total_leads > 0 else 0.0
    overall_cr = (total_conv / total_leads) if total_leads > 0 else 0.0

    y = h - 4.0*cm
    c.setFont("Helvetica-Bold", 12)
    c.drawString(2*cm, y, "Summary KPIs")
    y -= 0.8*cm

    c.setFont("Helvetica", 11)
    c.drawString(2*cm, y, f"Total Spend: £{total_spend:,.2f}")
    c.drawString(9*cm, y, f"Total Leads: {total_leads:,}")
    y -= 0.6*cm
    c.drawString(2*cm, y, f"Total Impressions: {total_impr:,}")
    c.drawString(9*cm, y, f"Overall CPL: £{overall_cpl:,.2f}")
    y -= 0.6*cm
    c.drawString(2*cm, y, f"Total Converted Leads: {total_conv:,}")
    c.drawString(9*cm, y, f"Overall Conversion Rate: {overall_cr*100:,.2f}%")

    c.showPage()
    c.save()
    buf.seek(0)
    return buf.read()


st.title("Marketing Performance Dashboard (Matplotlib)")
st.caption("Static-chart version — same logic, separate file.")

if "mode" not in st.session_state:
    st.session_state.mode = None

left, mid, right = st.columns([1, 3, 1])
with mid:
    c1, c2 = st.columns(2)
    if c1.button("LEADS", use_container_width=True):
        st.session_state.mode = "LEADS"
    if c2.button("MESSAGES", use_container_width=True):
        st.session_state.mode = "MESSAGES"

st.markdown('<div class="mode-note">Click <b>LEADS</b> or <b>MESSAGES</b> to continue.</div>', unsafe_allow_html=True)

if not st.session_state.mode:
    st.stop()

st.divider()

st.subheader("Upload Data")
uploaded_file = st.file_uploader("Upload Excel / CSV", type=["xlsx", "csv"])

if not uploaded_file:
    st.info("Upload an Excel/CSV file to continue.")
    st.stop()

try:
    if uploaded_file.name.lower().endswith(".csv"):
        df = pd.read_csv(uploaded_file)
    else:
        df = pd.read_excel(uploaded_file)

    df = normalize_columns(df)
    df = compute_metrics(df)
except Exception as e:
    st.error(f"Could not read your file. Reason: {e}")
    st.stop()

if st.session_state.mode == "MESSAGES":
    st.subheader("Messages Dashboard")
    st.info("Messages dashboard coming soon")
    st.stop()

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

# Editable table
st.markdown('<div class="card">', unsafe_allow_html=True)
st.markdown("### Editable Table (Impressions + Converted Leads)")

display_cols = [
    "month","brand","destination","impressions","cpl","spent_gbp","leads","converted_leads","conversion_rate"
]
table = d[display_cols].copy()
table["conversion_rate"] = table["conversion_rate"] * 100

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

d["impressions"] = pd.to_numeric(edited["impressions"], errors="coerce").fillna(0).astype(int)
d["converted_leads"] = pd.to_numeric(edited["converted_leads"], errors="coerce").fillna(0).astype(int)
d = compute_metrics(d)

st.markdown("</div>", unsafe_allow_html=True)

st.divider()

# KPIs
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

# Matplotlib charts (simple, clean)
st.subheader("Brand Level (Matplotlib)")

brand_agg = d.groupby("brand", as_index=False).agg(spend=("spent_gbp","sum"), leads=("leads","sum"))
brand_agg = brand_agg.sort_values("spend", ascending=True)

fig = plt.figure()
plt.barh(brand_agg["brand"], brand_agg["spend"])
plt.title("Spend by Brand")
plt.xlabel("Spend (GBP)")
st.pyplot(fig, clear_figure=True)

fig = plt.figure()
plt.barh(brand_agg["brand"], brand_agg["leads"])
plt.title("Leads by Brand")
plt.xlabel("Leads")
st.pyplot(fig, clear_figure=True)

st.divider()

# Exports
st.download_button(
    "Download filtered data (CSV)",
    d.to_csv(index=False).encode("utf-8"),
    file_name=f"filtered_{month}_{brand}_{destination}.csv".replace(" ", "_"),
    mime="text/csv",
)

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
