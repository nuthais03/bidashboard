import pandas as pd
import streamlit as st
import matplotlib.pyplot as plt

# --------------------------------
# Page setup
# --------------------------------
st.set_page_config(page_title="Performance Dashboard - Matplotlib", layout="wide")

st.title("Performance Dashboard - Matplotlib")
st.caption("Leads and Messages dashboards using Matplotlib")

# --------------------------------
# Session state
# --------------------------------
if "mode" not in st.session_state:
    st.session_state.mode = None   # Nothing selected initially

# --------------------------------
# Centered Buttons
# --------------------------------
st.markdown("### Select Dashboard")

left, center, right = st.columns([1,2,1])

with center:
    b1, b2 = st.columns(2)
    
    if b1.button("LEADS", use_container_width=True):
        st.session_state.mode = "Leads"
        
    if b2.button("MESSAGES", use_container_width=True):
        st.session_state.mode = "Messages"

st.divider()

# --------------------------------
# If nothing selected → stop
# --------------------------------
if st.session_state.mode is None:
    st.info("Select LEADS or MESSAGES to continue.")
    st.stop()

# --------------------------------
# Upload section (only after click)
# --------------------------------
st.subheader(f"{st.session_state.mode} Dashboard")
uploaded_file = st.file_uploader(
    f"Upload Excel file for {st.session_state.mode}",
    type=["xlsx"]
)

if not uploaded_file:
    st.info("Upload a file to view the dashboard.")
    st.stop()

# --------------------------------
# Load data
# --------------------------------
df = pd.read_excel(uploaded_file)
df.columns = df.columns.str.strip()

df = df.rename(columns={
    "Brand": "brand",
    "Destination": "destination",
    "Leads": "leads",
    "Spent (GBP)": "spent",
    "Month": "month"
})

required = {"brand", "destination", "leads", "spent", "month"}
missing = required - set(df.columns)

if missing:
    st.error(f"Missing columns: {', '.join(missing)}")
    st.stop()

df["leads"] = pd.to_numeric(df["leads"], errors="coerce").fillna(0)
df["spent"] = pd.to_numeric(df["spent"], errors="coerce").fillna(0)

# --------------------------------
# Filters
# --------------------------------
st.sidebar.header("Filters")

month = st.sidebar.selectbox("Month", sorted(df["month"].astype(str).unique()))
d = df[df["month"].astype(str) == month]

brand = st.sidebar.selectbox("Brand", ["All"] + sorted(d["brand"].unique()))
if brand != "All":
    d = d[d["brand"] == brand]

destination = st.sidebar.selectbox("Destination", ["All"] + sorted(d["destination"].unique()))
if destination != "All":
    d = d[d["destination"] == destination]

# --------------------------------
# LEADS Dashboard
# --------------------------------
if st.session_state.mode == "Leads":

    total_spend = d["spent"].sum()
    total_leads = d["leads"].sum()

    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Total Spend", f"£{total_spend:,.2f}")
    k2.metric("Total Leads", int(total_leads))
    k3.metric("Brands", d["brand"].nunique())
    k4.metric("Destinations", d["destination"].nunique())

    st.divider()

    col1, col2 = st.columns(2)

    # Spend by Brand
    spend_brand = d.groupby("brand")["spent"].sum().sort_values()
    fig1, ax1 = plt.subplots()
    spend_brand.plot(kind="barh", ax=ax1)
    ax1.set_title("Spend by Brand")
    col1.pyplot(fig1)

    # Leads by Brand
    leads_brand = d.groupby("brand")["leads"].sum().sort_values()
    fig2, ax2 = plt.subplots()
    leads_brand.plot(kind="barh", ax=ax2)
    ax2.set_title("Leads by Brand")
    col2.pyplot(fig2)

# --------------------------------
# MESSAGES Dashboard (placeholder)
# --------------------------------
else:
    st.info("Messages dashboard will be added later.")
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Messages", "-")
    m2.metric("Spend", "-")
    m3.metric("Brands", "-")
    m4.metric("Destinations", "-")
