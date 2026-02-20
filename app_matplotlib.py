import io
import pandas as pd
import plotly.express as px
import plotly.io as pio
import streamlit as st
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas

# -----------------------------
# Config
# -----------------------------
st.set_page_config(page_title="Performance Dashboard", layout="wide")
pio.templates.default = "plotly_dark"

MONTH_ORDER = [
    "January","February","March","April","May","June",
    "July","August","September","October","November","December"
]

RENAME_MAP = {
    "Brand": "brand",
    "Destination": "destination",
    "Leads": "leads",
    "Spent (GBP)": "spent_gbp",
    "Month": "month",
    "Impressions": "impressions",
    "CPL": "cpl",
    "Converted Leads": "converted_leads",
    "Conversion Rate": "conversion_rate",
    # later
    "Messages": "messages",
    "Spent Messages (GBP)": "spent_messages_gbp",
}

# -----------------------------
# UI polish (layout)
# -----------------------------
st.markdown(
    """
    <style>
      .block-container { padding-top: 1.1rem; padding-bottom: 2rem; max-width: 1200px; }
      section[data-testid="stSidebar"] { padding-top: 0.8rem; }

      /* Header */
      .title-wrap { margin-bottom: 0.2rem; }
      .subtitle { opacity: 0.8; margin-top: -0.2rem; }

      /* Center mode buttons */
      .mode-wrap { display: flex; justify-content: center; margin: 1rem 0 0.8rem 0; }
      .mode-card {
        width: 100%;
        border: 1px solid rgba(255,255,255,0.10);
        border-radius: 14px;
        padding: 18px 18px 6px 18px;
        background: rgba(255,255,255,0.03);
      }

      /* Section cards */
      .card {
        border: 1px solid rgba(255,255,255,0.10);
        border-radius: 14px;
        padding: 18px;
        background: rgba(255,255,255,0.03);
        margin-top: 14px;
      }

      .card-title { font-size: 1.1rem; font-weight: 700; margin-bottom: 0.6rem; }
      .muted { opacity: 0.8; font-size: 0.92rem; }

      /* Metrics */
      [data-testid="stMetricValue"] { font-size: 1.6rem; }
      [data-testid="stMetricLabel"] { opacity: 0.85; }

      /* Remove excessive gaps */
      .stDivider { margin-top: 0.7rem; margin-bottom: 0.7rem; }
    </style>
    """,
    unsafe_allow_html=True,
)

# -----------------------------
# Helpers
# -----------------------------
def ensure_columns(df: pd.DataFrame, mode: str) -> pd.DataFrame:
    for col in ["brand", "destination", "month"]:
        if col not in df.columns:
            df[col] = ""

    if mode == "LEADS":
        for col in ["leads", "spent_gbp", "impressions", "cpl", "converted_leads", "conversion_rate"]:
            if col not in df.columns:
                df[col] = 0
    else:  # MESSAGES placeholder
        for col in ["messages", "spent_messages_gbp"]:
            if col not in df.columns:
                df[col] = 0

    return df

def clean_types(df: pd.DataFrame) -> pd.DataFrame:
    df["month"] = df["month"].astype(str).str.strip()
    df["brand"] = df["brand"].astype(str).str.strip()
    df["destination"] = df["destination"].astype(str).str.strip()
    df["month"] = pd.Categorical(df["month"], categories=MONTH_ORDER, ordered=True)

    for c in ["leads", "impressions", "converted_leads", "messages"]:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0).astype(int)

    for c in ["spent_gbp", "spent_messages_gbp", "cpl", "conversion_rate"]:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0.0)

    return df

def compute_derived(d: pd.DataFrame) -> pd.DataFrame:
    out = d.copy()
    leads_safe = out["leads"].replace(0, pd.NA)
    out["cpl"] = (out["spent_gbp"] / leads_safe).fillna(0.0)
    out["conversion_rate"] = (out["converted_leads"] / leads_safe).fillna(0.0)
    return out

def pdf_summary(mode: str, month: str, brand: str, destination: str, kpis: dict) -> bytes:
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    width, height = A4
    y = height - 60

    c.setFont("Helvetica-Bold", 16)
    c.drawString(50, y, f"Performance Summary ({mode})")
    y -= 20

    c.setFont("Helvetica", 10)
    c.drawString(50, y, f"Month={month} | Brand={brand} | Destination={destination}")
    y -= 30

    c.setFont("Helvetica-Bold", 12)
    c.drawString(50, y, "KPIs")
    y -= 18

    c.setFont("Helvetica", 10)
    for k, v in kpis.items():
        c.drawString(60, y, f"{k}: {v}")
        y -= 14

    c.showPage()
    c.save()
    buf.seek(0)
    return buf.read()

# -----------------------------
# Header
# -----------------------------
st.markdown('<div class="title-wrap">', unsafe_allow_html=True)
st.title("Performance Dashboard")
st.markdown('<div class="subtitle">LEADS and MESSAGES dashboards with upload → filters → insights → export.</div>', unsafe_allow_html=True)
st.markdown('</div>', unsafe_allow_html=True)

# -----------------------------
# Mode selector (centered)
# -----------------------------
if "mode" not in st.session_state:
    st.session_state.mode = None

st.markdown('<div class="mode-card">', unsafe_allow_html=True)
st.markdown("### Select Dashboard")
b1, b2, b3 = st.columns([1, 2, 2])
with b2:
    if st.button("✅ LEADS", use_container_width=True):
        st.session_state.mode = "LEADS"
with b3:
    if st.button("💬 MESSAGES", use_container_width=True):
        st.session_state.mode = "MESSAGES"
st.markdown(f"<div class='muted'><b>Current:</b> {st.session_state.mode if st.session_state.mode else 'Not selected'}</div>", unsafe_allow_html=True)
st.markdown('</div>', unsafe_allow_html=True)

mode = st.session_state.mode
if not mode:
    st.info("Select **LEADS** or **MESSAGES** to continue.")
    st.stop()

# -----------------------------
# Upload (inside a card)
# -----------------------------
st.markdown('<div class="card">', unsafe_allow_html=True)
st.markdown('<div class="card-title">Upload Data</div>', unsafe_allow_html=True)
uploaded_file = st.file_uploader("Upload Excel (.xlsx)", type=["xlsx"], label_visibility="collapsed")
st.markdown('<div class="muted">Tip: Keep your file private — upload stays only inside this app session.</div>', unsafe_allow_html=True)
st.markdown('</div>', unsafe_allow_html=True)

if not uploaded_file:
    st.warning("Upload your Excel file to view the dashboard.")
    st.stop()

# -----------------------------
# Load
# -----------------------------
df = pd.read_excel(uploaded_file)
df.columns = df.columns.str.strip()
df = df.rename(columns=RENAME_MAP)
df = ensure_columns(df, mode)
df = clean_types(df)

# Sidebar filters AFTER upload
st.sidebar.header("Filters")
available_months = [m for m in MONTH_ORDER if m in df["month"].dropna().unique().tolist()]
if not available_months:
    available_months = sorted(df["month"].dropna().astype(str).unique().tolist())

month = st.sidebar.selectbox("Month", available_months)
d = df[df["month"] == month].copy()

brand = st.sidebar.selectbox("Brand", ["All"] + sorted(d["brand"].dropna().unique()))
if brand != "All":
    d = d[d["brand"] == brand]

destination = st.sidebar.selectbox("Destination", ["All"] + sorted(d["destination"].dropna().unique()))
if destination != "All":
    d = d[d["destination"] == destination]

# -----------------------------
# LEADS dashboard
# -----------------------------
if mode == "LEADS":
    # Manual inputs
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown('<div class="card-title">Manual Inputs (Optional)</div>', unsafe_allow_html=True)
    st.markdown('<div class="muted">Edit <b>Converted Leads</b> and <b>Impressions</b>. Conversion Rate updates automatically.</div>', unsafe_allow_html=True)

    edit_cols = ["brand", "destination", "month", "leads", "spent_gbp", "impressions", "converted_leads"]
    edited = st.data_editor(
        d[edit_cols].copy(),
        use_container_width=True,
        num_rows="fixed",
        column_config={
            "converted_leads": st.column_config.NumberColumn("Converted Leads", min_value=0, step=1),
            "impressions": st.column_config.NumberColumn("Impressions", min_value=0, step=1),
        }
    )
    d["converted_leads"] = edited["converted_leads"].values
    d["impressions"] = edited["impressions"].values
    st.markdown('</div>', unsafe_allow_html=True)

    # Derived metrics
    d = compute_derived(d)

    # KPIs card
    total_spend = float(d["spent_gbp"].sum())
    total_leads = int(d["leads"].sum())
    total_impr = int(d["impressions"].sum())
    total_conv = int(d["converted_leads"].sum())
    cpl_overall = (total_spend / total_leads) if total_leads else 0.0
    conv_rate_overall = (total_conv / total_leads) if total_leads else 0.0

    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown('<div class="card-title">Summary</div>', unsafe_allow_html=True)
    k1, k2, k3, k4, k5, k6 = st.columns(6)
    k1.metric("Spend", f"£{total_spend:,.2f}")
    k2.metric("Leads", f"{total_leads:,}")
    k3.metric("Impressions", f"{total_impr:,}")
    k4.metric("CPL", f"£{cpl_overall:,.2f}")
    k5.metric("Converted", f"{total_conv:,}")
    k6.metric("Conv. Rate", f"{conv_rate_overall*100:,.2f}%")
    st.markdown('</div>', unsafe_allow_html=True)

    # Charts card
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown('<div class="card-title">Brand Performance</div>', unsafe_allow_html=True)
    c1, c2 = st.columns(2)

    spend_by_brand = d.groupby("brand", as_index=False)["spent_gbp"].sum().sort_values("spent_gbp", ascending=True)
    fig_spend = px.bar(spend_by_brand, x="spent_gbp", y="brand", orientation="h", title="Spend by Brand")
    fig_spend.update_layout(xaxis_title="Spend (GBP)", yaxis_title="Brand")
    c1.plotly_chart(fig_spend, use_container_width=True)

    leads_by_brand = d.groupby("brand", as_index=False)["leads"].sum().sort_values("leads", ascending=True)
    fig_leads = px.bar(leads_by_brand, x="leads", y="brand", orientation="h", title="Leads by Brand")
    fig_leads.update_layout(xaxis_title="Leads", yaxis_title="Brand")
    c2.plotly_chart(fig_leads, use_container_width=True)
    st.markdown('</div>', unsafe_allow_html=True)

    # Decomposition card (tables only, clean)
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown('<div class="card-title">Decomposition: Brand → Destination</div>', unsafe_allow_html=True)
    view = st.radio("Breakdown level", ["Brand summary", "Destination summary", "Brand → destination detail"], horizontal=True)

    if view == "Brand summary":
        table = d.groupby("brand", as_index=False).agg(
            spent_gbp=("spent_gbp","sum"),
            leads=("leads","sum"),
            impressions=("impressions","sum"),
            converted_leads=("converted_leads","sum"),
        )
    elif view == "Destination summary":
        table = d.groupby("destination", as_index=False).agg(
            spent_gbp=("spent_gbp","sum"),
            leads=("leads","sum"),
            impressions=("impressions","sum"),
            converted_leads=("converted_leads","sum"),
        )
    else:
        table = d.groupby(["brand","destination"], as_index=False).agg(
            spent_gbp=("spent_gbp","sum"),
            leads=("leads","sum"),
            impressions=("impressions","sum"),
            converted_leads=("converted_leads","sum"),
        )

    table["cpl"] = table.apply(lambda r: (r["spent_gbp"]/r["leads"]) if r["leads"] else 0.0, axis=1)
    table["conversion_rate"] = table.apply(lambda r: (r["converted_leads"]/r["leads"]) if r["leads"] else 0.0, axis=1)
    table = table.sort_values("spent_gbp", ascending=False)

    st.dataframe(table, use_container_width=True)
    st.markdown('</div>', unsafe_allow_html=True)

    # Export card
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown('<div class="card-title">Export</div>', unsafe_allow_html=True)

    st.download_button(
        "Download filtered data (CSV)",
        d.to_csv(index=False).encode("utf-8"),
        file_name=f"filtered_LEADS_{month}_{brand}_{destination}.csv".replace(" ", "_"),
        mime="text/csv"
    )

    pdf_bytes = pdf_summary(
        mode="LEADS",
        month=month,
        brand=brand,
        destination=destination,
        kpis={
            "Total Spend (GBP)": f"£{total_spend:,.2f}",
            "Total Leads": f"{total_leads:,}",
            "Impressions": f"{total_impr:,}",
            "CPL": f"£{cpl_overall:,.2f}",
            "Converted Leads": f"{total_conv:,}",
            "Conversion Rate": f"{conv_rate_overall*100:,.2f}%"
        }
    )
    st.download_button(
        "Download PDF Summary",
        data=pdf_bytes,
        file_name=f"summary_LEADS_{month}_{brand}_{destination}.pdf".replace(" ", "_"),
        mime="application/pdf"
    )
    st.markdown('</div>', unsafe_allow_html=True)

    with st.expander("Show full filtered data"):
        st.dataframe(d, use_container_width=True)

# -----------------------------
# MESSAGES placeholder
# -----------------------------
else:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown('<div class="card-title">MESSAGES Dashboard (Coming Soon)</div>', unsafe_allow_html=True)
    st.info("This space is reserved. When you add Messages columns in your Excel, we will activate charts + KPIs here.")
    st.markdown('</div>', unsafe_allow_html=True)
