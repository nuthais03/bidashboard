import pandas as pd
import streamlit as st
import matplotlib.pyplot as plt

# --------------------------------
# Page setup
# --------------------------------
st.set_page_config(page_title="Performance Dashboard (Matplotlib)", layout="wide")

st.title("Performance Dashboard - Matplotlib")
st.caption("Leads and Messages dashboards using Matplotlib.")

# --------------------------------
# Dashboard switch
# --------------------------------
col1, col2 = st.columns(2)

if "mode" not in st.session_state:
    st.session_state.mode = "Leads"

with col1:
    if st.button("LEADS", use_container_width=True):
        st.session_state.mode = "Leads"

with col2:
    if st.button("MESSAGES", use_container_width=True):
        st.session_state.mode = "Messages"

st.write(f"**Current Dashboard:** {st.session_state.mode}")
st.divider()

# --------------------------------
# Upload file
# --------------------------------
st.subheader("Upload Data")
uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx"])

if not uploaded_file:
    st.info("Upload an Excel file to continue.")
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
# LEADS DASHBOARD
# --------------------------------
if st.session_state.mode == "Leads":

    st.subheader("Leads Dashboard")

    total_spend = d["spent"].sum()
    total_leads = d["leads"].sum()

    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Total Spend", f"£{total_spend:,.2f}")
    k2.metric("Total Leads", int(total_leads))
    k3.metric("Brands", d["brand"].nunique())
    k4.metric("Destinations", d["destination"].nunique())

    st.divider()

    # --------------------------------
    # Matplotlib Charts
    # --------------------------------
    colA, colB = st.columns(2)

    # Spend by Brand
    spend_brand = d.groupby("brand")["spent"].sum().sort_values()

    fig1, ax1 = plt.subplots()
    spend_brand.plot(kind="barh", ax=ax1)
    ax1.set_title("Spend by Brand")
    ax1.set_xlabel("GBP")
    ax1.set_ylabel("Brand")
    colA.pyplot(fig1)

    # Leads by Brand
    leads_brand = d.groupby("brand")["leads"].sum().sort_values()

    fig2, ax2 = plt.subplots()
    leads_brand.plot(kind="barh", ax=ax2)
    ax2.set_title("Leads by Brand")
    ax2.set_xlabel("Leads")
    ax2.set_ylabel("Brand")
    colB.pyplot(fig2)

    st.divider()

    # Destination chart
    dest_data = d.groupby("destination")[["spent", "leads"]].sum()

    fig3, ax3 = plt.subplots()
    dest_data["spent"].sort_values().plot(kind="barh", ax=ax3)
    ax3.set_title("Spend by Destination")
    st.pyplot(fig3)

# --------------------------------
# MESSAGES DASHBOARD (placeholder)
# --------------------------------
else:
    st.subheader("Messages Dashboard")
    st.info("This section is reserved for Messages campaign data.")

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Total Messages", "-")
    m2.metric("Total Spend", "-")
    m3.metric("Brands", "-")
    m4.metric("Destinations", "-")

    st.write("Charts will be added once Messages data is available.")
