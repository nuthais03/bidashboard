import io
import numpy as np
import pandas as pd
import streamlit as st

import plotly.express as px
import plotly.io as pio

# PDF
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.units import cm

# --------------------------------
# Page setup
# --------------------------------
st.set_page_config(page_title="Performance Dashboard", layout="wide")
pio.templates.default = "plotly_dark"

# --------------------------------
# Styling
# --------------------------------
st.markdown(
    """
    <style>
      .block-container { padding-top: 1.2rem; padding-bottom: 2rem; }
      section[data-testid="stSidebar"] { padding-top: 1rem; }
      h1, h2, h3 { letter-spacing: 0.2px; }
      [data-testid="stCaptionContainer"] { opacity: 0.85; }
      [data-testid="stMetricValue"] { font-size: 1.7rem; }
      .mode-wrap { display:flex; justify-content:center; gap:14px; margin: 18px 0 8px 0; }
      .mode-note { text-align:center; opacity:0.85; margin-bottom: 12px; }
      .card { border:1px solid rgba(255,255,255,0.10); border-radius:14px; padding:14px; background:rgba(255,255,255,0.02); }
      .muted { opacity:0.85; }
    </style>
    """,
    unsafe_allow_html=True,
)

# --------------------------------
# Helpers
# --------------------------------
MONTH_ORDER = [
    "January","February","March","April","May","June",
    "July","August","September","October","November","December"
]

def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Accepts either your new template columns, or old template with Spent (GBP)."""
    df = df.copy()
    df.columns = df.columns.str.strip()

    # Standard rename map (flexible)
    rename_map = {
        "Month": "month",
        "Brand": "brand",
        "Destination": "destination",
        "Impressions": "impressions",
        "CPL": "cpl",
        "Leads": "leads",
        "Converted Leads": "converted_leads",
        "Conversion Rate": "conversion_rate",
        "Spent (GBP)": "spent_gbp",  # legacy support
        "Spent": "spent_gbp",
    }
    df = df.rename(columns={k: v for k, v in rename_map.items() if k in df.columns})

    # Ensure core cols exist
    required_any = {"month", "brand", "destination", "leads"}
    missing_any = required_any - set(df.columns)
    if missing_any:
        raise ValueError(f"Missing required columns: {', '.join(sorted(missing_any))}")

    # Optional cols
    if "impressions" not in df.columns:
        df["impressions"] = 0
    if "converted_leads" not in df.columns:
        df["converted_leads"] = 0
    if "cpl" not in df.columns:
        df["cpl"] = np.nan
    if "conversion_rate" not in df.columns:
        df["conversion_rate"] = np.nan
    if "spent_gbp" not in df.columns:
        df["spent_gbp"] = np.nan  # only needed for computing CPL if CPL missing

    # Clean types
    df["month"] = df["month"].astype(str).str.strip()
    df["brand"] = df["brand"].astype(str).str.strip()
    df["destination"] = df["destination"].astype(str).str.strip()

    df["leads"] = pd.to_numeric(df["leads"], errors="coerce").fillna(0).astype(int)
    df["impressions"] = pd.to_numeric(df["impressions"], errors="coerce").fillna(0).astype(int)
    df["converted_leads"] = pd.to_numeric(df["converted_leads"], errors="coerce").fillna(0).astype(int)

    df["cpl"] = pd.to_numeric(df["cpl"], errors="coerce")
    df["conversion_rate"] = pd.to_numeric(df["conversion_rate"], errors="coerce")
    df["spent_gbp"] = pd.to_numeric(df["spent_gbp"], errors="coerce")

    # Month order
    if df["month"].isin(MONTH_ORDER).any():
        df["month"] = pd.Categorical(df["month"], categories=MONTH_ORDER, ordered=True)

    return df


def compute_row_metrics(df: pd.DataFrame) -> pd.DataFrame:
    """Compute CPL and conversion rate row-wise if missing."""
    out = df.copy()

    leads_safe = out["leads"].replace(0, np.nan)

    # Conversion rate (always computed from converted/leads)
    out["conversion_rate"] = (out["converted_leads"] / leads_safe).fillna(0.0)

    # CPL:
    # If CPL provided -> use it
    # Else if spent_gbp exists -> compute spent/leads
    # Else -> 0
    cpl_from_spend = (out["spent_gbp"] / leads_safe)
    out["cpl"] = out["cpl"].where(out["cpl"].notna(), cpl_from_spend).fillna(0.0)

    return out


def build_pdf_report(mode: str, filters: dict, d: pd.DataFrame) -> bytes:
    """Simple professional PDF summary."""
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    w, h = A4

    # Header
    c.setFont("Helvetica-Bold", 16)
    c.drawString(2*cm, h - 2*cm, f"Performance Report — {mode}")

    c.setFont("Helvetica", 10)
    c.drawString(2*cm, h - 2.8*cm, f"Month: {filters.get('month','All')} | Brand: {filters.get('brand','All')} | Destination: {filters.get('destination','All')}")

    # KPIs
    total_leads = int(d["leads"].sum())
    total_converted = int(d["converted_leads"].sum())
    total_impr = int(d["impressions"].sum())

    # overall CPL: weighted avg by leads if possible; else mean
    if total_leads > 0:
        overall_cpl = float((d["cpl"] * d["leads"]).sum() / total_leads)
        overall_cr = float(total_converted / total_leads)
    else:
        overall_cpl = float(d["cpl"].mean()) if len(d) else 0.0
        overall_cr = 0.0

    y = h - 4.0*cm
    c.setFont("Helvetica-Bold", 12)
    c.drawString(2*cm, y, "Summary KPIs")
    y -= 0.7*cm

    c.setFont("Helvetica", 11)
    c.drawString(2*cm, y, f"Leads: {total_leads:,}")
    c.drawString(8*cm, y, f"Converted Leads: {total_converted:,}")
    y -= 0.6*cm
    c.drawString(2*cm, y, f"Impressions: {total_impr:,}")
    c.drawString(8*cm, y, f"Overall CPL: £{overall_cpl:,.2f}")
    y -= 0.6*cm
    c.drawString(2*cm, y, f"Overall Conversion Rate: {overall_cr*100:,.2f}%")

    # Top 10 table (by leads)
    y -= 1.2*cm
    c.setFont("Helvetica-Bold", 12)
    c.drawString(2*cm, y, "Top Rows (by Leads)")
    y -= 0.7*cm

    table = d.copy()
    table["conversion_rate_pct"] = table["conversion_rate"] * 100
    top = table.sort_values("leads", ascending=False).head(10)

    c.setFont("Helvetica-Bold", 9)
    c.drawString(2*cm, y, "Brand")
    c.drawString(6.3*cm, y, "Destination")
    c.drawString(11.0*cm, y, "Leads")
    c.drawString(13.0*cm, y, "CPL")
    c.drawString(15.0*cm, y, "CR%")
    y -= 0.5*cm

    c.setFont("Helvetica", 9)
    for _, r in top.iterrows():
        if y < 2.5*cm:
            c.showPage()
            y = h - 2.0*cm
            c.setFont("Helvetica", 9)

        c.drawString(2*cm, y, str(r["brand"])[:18])
        c.drawString(6.3*cm, y, str(r["destination"])[:20])
        c.drawRightString(12.4*cm, y, f"{int(r['leads']):,}")
        c.drawRightString(14.6*cm, y, f"£{float(r['cpl']):,.2f}")
        c.drawRightString(19.0*cm, y, f"{float(r['conversion_rate_pct']):.2f}%")
        y -= 0.45*cm

    c.showPage()
    c.save()
    buf.seek(0)
    return buf.read()


# --------------------------------
# Title + mode selector
# --------------------------------
st.title("Performance Dashboard")
st.caption("Choose a dashboard (Leads / Messages), upload data, and explore insights.")

if "mode" not in st.session_state:
    st.session_state.mode = None

# Center buttons
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

# --------------------------------
# Upload section (shown after selecting mode)
# --------------------------------
st.subheader("Upload Data")
uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx"])

if not uploaded_file:
    st.info("Upload an Excel file to continue.")
    st.stop()

# --------------------------------
# Load data
# --------------------------------
try:
    df = pd.read_excel(uploaded_file)
    df = normalize_columns(df)
    df = compute_row_metrics(df)
except Exception as e:
    st.error(f"Could not read your file. Reason: {e}")
    st.stop()

# --------------------------------
# Filters
# --------------------------------
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

# Recompute after filtering (safe)
d = compute_row_metrics(d)

# --------------------------------
# Messages mode placeholder
# --------------------------------
if st.session_state.mode == "MESSAGES":
    st.subheader("Messages Dashboard (Placeholder)")
    st.write("This section is reserved for your Messages campaign dashboard.")
    st.info("Later, we’ll connect the Messages export columns and build similar KPIs/charts.")
    st.stop()

# --------------------------------
# Export filtered CSV
# --------------------------------
st.download_button(
    "Download filtered data (CSV)",
    d.to_csv(index=False).encode("utf-8"),
    file_name=f"filtered_{month}_{brand}_{destination}.csv".replace(" ", "_"),
    mime="text/csv"
)

# --------------------------------
# --------------------------------
# Manual Inputs table (edit only inputs)
# --------------------------------
st.markdown('<div class="card">', unsafe_allow_html=True)
st.markdown("### Manual Inputs (Optional)")
st.markdown(
    '<div class="muted">Edit <b>Impressions</b> and <b>Converted Leads</b>. '
    '<b>CPL</b> and <b>Conversion Rate</b> update automatically below.</div>',
    unsafe_allow_html=True
)

input_cols = ["month","brand","destination","impressions","leads","converted_leads"]
inputs = d[input_cols].copy()

edited = st.data_editor(
    inputs,
    use_container_width=True,
    hide_index=True,
    column_config={
        "month": st.column_config.TextColumn("Month", disabled=True),
        "brand": st.column_config.TextColumn("Brand", disabled=True),
        "destination": st.column_config.TextColumn("Destination", disabled=True),
        "impressions": st.column_config.NumberColumn("Impressions", min_value=0, step=1),
        "leads": st.column_config.NumberColumn("Leads", disabled=True),
        "converted_leads": st.column_config.NumberColumn("Converted Leads", min_value=0, step=1),
    },
)

# Push edits back + compute metrics
d["impressions"] = pd.to_numeric(edited["impressions"], errors="coerce").fillna(0).astype(int)
d["converted_leads"] = pd.to_numeric(edited["converted_leads"], errors="coerce").fillna(0).astype(int)
d = compute_row_metrics(d)

# Show calculated result table (THIS will update correctly)
result = d[["month","brand","destination","impressions","cpl","leads","converted_leads","conversion_rate"]].copy()
result["conversion_rate"] = result["conversion_rate"] * 100  # convert to %
result = result.rename(columns={"conversion_rate": "conversion_rate_%"})

st.dataframe(
    result,
    use_container_width=True,
    hide_index=True,
)

st.markdown("</div>", unsafe_allow_html=True)

# push edits back
d["impressions"] = pd.to_numeric(edited["impressions"], errors="coerce").fillna(0).astype(int)
d["converted_leads"] = pd.to_numeric(edited["converted_leads"], errors="coerce").fillna(0).astype(int)
d = compute_row_metrics(d)

st.markdown("</div>", unsafe_allow_html=True)

st.divider()

# --------------------------------
# Summary KPIs
# --------------------------------
total_leads = int(d["leads"].sum())
total_converted = int(d["converted_leads"].sum())
total_impr = int(d["impressions"].sum())

if total_leads > 0:
    overall_cpl = float((d["cpl"] * d["leads"]).sum() / total_leads)
    overall_cr = float(total_converted / total_leads)
else:
    overall_cpl = float(d["cpl"].mean()) if len(d) else 0.0
    overall_cr = 0.0

k1, k2, k3, k4, k5, k6 = st.columns(6)
k1.metric("Spend (optional)", "—" if d["spent_gbp"].isna().all() else f"£{float(d['spent_gbp'].sum()):,.2f}")
k2.metric("Leads", f"{total_leads:,}")
k3.metric("Impressions", f"{total_impr:,}")
k4.metric("CPL (Overall)", f"£{overall_cpl:,.2f}")
k5.metric("Converted", f"{total_converted:,}")
k6.metric("Conv. Rate (Overall)", f"{overall_cr*100:,.2f}%")

st.divider()

# --------------------------------
# Charts
# --------------------------------
st.subheader("Brand Performance")

c1, c2 = st.columns(2)

brand_leads = d.groupby("brand", as_index=False)["leads"].sum().sort_values("leads", ascending=True)
fig1 = px.bar(brand_leads, x="leads", y="brand", orientation="h", title="Leads by Brand")
c1.plotly_chart(fig1, use_container_width=True)

brand_cpl = d.groupby("brand", as_index=False).apply(
    lambda x: pd.Series({
        "cpl": (x["cpl"] * x["leads"]).sum() / x["leads"].sum() if x["leads"].sum() > 0 else 0
    })
).reset_index()
brand_cpl = brand_cpl.sort_values("cpl", ascending=True)
fig2 = px.bar(brand_cpl, x="cpl", y="brand", orientation="h", title="CPL by Brand (Weighted)")
fig2.update_layout(xaxis_title="CPL", yaxis_title="Brand")
c2.plotly_chart(fig2, use_container_width=True)

st.divider()

st.subheader("Destination Performance")

top_n = st.slider("Number of destinations to show", 5, 30, 10)

dest = (
    d.groupby("destination", as_index=False)
    .agg(leads=("leads", "sum"), converted=("converted_leads", "sum"))
)
dest["conversion_rate"] = np.where(dest["leads"] > 0, dest["converted"] / dest["leads"], 0.0)
dest = dest.sort_values("leads", ascending=False).head(top_n)

c3, c4 = st.columns(2)
fig3 = px.bar(dest.sort_values("leads"), x="leads", y="destination", orientation="h", title="Top Destinations by Leads")
c3.plotly_chart(fig3, use_container_width=True)

fig4 = px.bar(
    dest.sort_values("conversion_rate"),
    x="conversion_rate", y="destination",
    orientation="h",
    title="Top Destinations by Conversion Rate"
)
fig4.update_layout(xaxis_tickformat=".0%")
c4.plotly_chart(fig4, use_container_width=True)

st.divider()

# --------------------------------
# Report download (PDF)
# --------------------------------
filters = {"month": month, "brand": brand, "destination": destination}
pdf_bytes = build_pdf_report("LEADS", filters, d)

st.download_button(
    "Download Summary Report (PDF)",
    data=pdf_bytes,
    file_name=f"report_{month}_{brand}_{destination}.pdf".replace(" ", "_"),
    mime="application/pdf"
)

# --------------------------------
# Full table
# --------------------------------
with st.expander("Show full filtered table"):
    show_cols = [
        "month","brand","destination","impressions","cpl","leads","converted_leads","conversion_rate"
    ]
    out = d[show_cols].copy()
    out["conversion_rate"] = out["conversion_rate"] * 100
    out = out.rename(columns={"conversion_rate": "conversion_rate_%"})
    st.dataframe(out, use_container_width=True)
