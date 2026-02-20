import pandas as pd
import plotly.express as px
import plotly.io as pio
import streamlit as st

# -----------------------------
# Page setup
# -----------------------------
st.set_page_config(page_title="Performance Dashboard", layout="wide")
pio.templates.default = "plotly_dark"

st.markdown(
    """
    <style>
      .block-container { padding-top: 1.2rem; padding-bottom: 2rem; }
      h1 { letter-spacing: 0.2px; }
      .center-row { display:flex; justify-content:center; gap:14px; margin-top: 8px; margin-bottom: 8px; }
      .center-row .stButton>button { width: 360px; height: 48px; font-weight: 700; }
    </style>
    """,
    unsafe_allow_html=True,
)

st.title("Performance Dashboard")
st.caption("Choose a dashboard (Leads / Messages) → upload data → analyze.")

# -----------------------------
# Session state
# -----------------------------
if "mode" not in st.session_state:
    st.session_state.mode = None

# -----------------------------
# Center buttons
# -----------------------------
st.markdown("### Select Dashboard")
c1, c2, c3 = st.columns([1, 2, 1])

with c2:
    b1, b2 = st.columns(2)
    if b1.button("LEADS", use_container_width=True):
        st.session_state.mode = "Leads"
    if b2.button("MESSAGES", use_container_width=True):
        st.session_state.mode = "Messages"

st.divider()

# -----------------------------
# Stop until user selects
# -----------------------------
if st.session_state.mode is None:
    st.info("Click **LEADS** or **MESSAGES** to continue.")
    st.stop()

st.subheader(f"Current Dashboard: {st.session_state.mode}")

# -----------------------------
# Upload only after selection
# -----------------------------
uploaded_file = st.file_uploader(
    f"Upload Excel file for {st.session_state.mode}",
    type=["xlsx"]
)

if not uploaded_file:
    st.info("Upload an Excel file to continue.")
    st.stop()

# -----------------------------
# Load & clean
# -----------------------------
df = pd.read_excel(uploaded_file)
df.columns = df.columns.str.strip()

# Required for Leads dashboard (same as your working file)
rename_map = {
    "Brand": "brand",
    "Destination": "destination",
    "Leads": "leads",
    "Spent (GBP)": "spent_gbp",
    "Month": "month"
}
df = df.rename(columns=rename_map)

required_cols = {"brand", "destination", "leads", "spent_gbp", "month"}
missing = required_cols - set(df.columns)

if st.session_state.mode == "Leads":
    if missing:
        st.error(f"Missing required columns: {', '.join(sorted(missing))}")
        st.stop()

    df["leads"] = pd.to_numeric(df["leads"], errors="coerce").fillna(0).astype(int)
    df["spent_gbp"] = pd.to_numeric(df["spent_gbp"], errors="coerce").fillna(0.0)

    df["month"] = df["month"].astype(str).str.strip()
    df["brand"] = df["brand"].astype(str).str.strip()
    df["destination"] = df["destination"].astype(str).str.strip()

    month_order = [
        "January","February","March","April","May","June",
        "July","August","September","October","November","December"
    ]
    df["month"] = pd.Categorical(df["month"], categories=month_order, ordered=True)

    # -----------------------------
    # Sidebar filters
    # -----------------------------
    st.sidebar.header("Filters")
    available_months = [m for m in month_order if m in df["month"].dropna().unique().tolist()]
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
    # KPIs
    # -----------------------------
    total_spend = float(d["spent_gbp"].sum())
    total_leads = int(d["leads"].sum())

    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Total Spend", f"£{total_spend:,.2f}")
    k2.metric("Total Leads", f"{total_leads:,}")
    k3.metric("Brands", f"{d['brand'].nunique():,}")
    k4.metric("Destinations", f"{d['destination'].nunique():,}")

    st.divider()

    # -----------------------------
    # Charts
    # -----------------------------
    c1, c2 = st.columns(2)

    spend_by_brand = d.groupby("brand", as_index=False)["spent_gbp"].sum().sort_values("spent_gbp", ascending=True)
    fig1 = px.bar(spend_by_brand, x="spent_gbp", y="brand", orientation="h", title="Spend by Brand")
    fig1.update_layout(xaxis_title="Spend (GBP)", yaxis_title="Brand")
    c1.plotly_chart(fig1, use_container_width=True)

    leads_by_brand = d.groupby("brand", as_index=False)["leads"].sum().sort_values("leads", ascending=True)
    fig2 = px.bar(leads_by_brand, x="leads", y="brand", orientation="h", title="Leads by Brand")
    fig2.update_layout(xaxis_title="Leads", yaxis_title="Brand")
    c2.plotly_chart(fig2, use_container_width=True)

    st.divider()

    # -----------------------------
    # Table
    # -----------------------------
    with st.expander("Show filtered data table"):
        st.dataframe(d, use_container_width=True)

else:
    # Messages dashboard placeholder (keeps space)
    st.warning("Messages dashboard is not connected yet. (Placeholder)")
    m1, m2, m3 = st.columns(3)
    m1.metric("Messages", "-")
    m2.metric("Spend", "-")
    m3.metric("CPL", "-")
