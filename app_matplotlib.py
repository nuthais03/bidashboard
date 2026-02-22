# app.py  ✅ Plotly / Production version (Streamlit Cloud ready)
# - LEADS + MESSAGES buttons (top-center)
# - Upload appears only after selection
# - Final columns order (as per your requirement)
# - Editable table: ONLY Impressions + Converted Leads
# - CPL = Spent / Leads
# - Conversion Rate = Converted / Leads
# - ONE Apply + ONE Reset (inside a form) -> one click updates conversion rate
# - Charts + Decomposition + CSV + PDF

import io
import numpy as np
import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.io as pio

# Optional PDF (works when reportlab is in requirements.txt)
PDF_AVAILABLE = True
try:
    from reportlab.lib.pagesizes import A4
    from reportlab.pdfgen import canvas
    from reportlab.lib.units import cm
except ModuleNotFoundError:
    PDF_AVAILABLE = False

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

# Internal column order (matches your final requirement)
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
            + "  |  Required: Month, Brand, Destination, Leads, Spent (GBP)"
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


def apply_overrides(base_df: pd.DataFrame, overrides: dict) -> pd.DataFrame:
    """Apply stored overrides for impressions & converted leads by row_id."""
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
# Header + Mode Selector
# -----------------------------
st.title("Marketing Performance Dashboard")
st.caption("Private performance dashboard — upload Excel/CSV and explore insights.")

if "mode" not in st.session_state:
    st.session_state.mode = None

# store overrides safely
if "overrides" not in st.session_state:
    st.session_state.overrides = {"impressions": {}, "converted_leads": {}}

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
# Upload (only after mode)
# -----------------------------
st.subheader("Upload Data")
uploaded_file = st.file_uploader("Upload Excel / CSV", type=["xlsx", "csv"])

if not uploaded_file:
    st.info("Upload an Excel/CSV file to continue.")
    st.stop()

# -----------------------------
# Load data
# -----------------------------
try:
    if uploaded_file.name.lower().endswith(".csv"):
        df = pd.read_csv(uploaded_file)
    else:
        df = pd.read_excel(uploaded_file)

    df = normalize_columns(df)
except Exception as e:
    st.error(f"Could not read your file. Reason: {e}")
    st.stop()

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
st.sidebar.header("Filters")

available_months = [m for m in MONTH_ORDER if m in df["month"].dropna().astype(str).unique().tolist()]
if not available_months:
    available_months = sorted(df["month"].dropna().astype(str).unique().tolist())

month = st.sidebar.selectbox("Month", available_months)
d0 = df[df["month"].astype(str) == str(month)].copy()

brand = st.sidebar.selectbox("Brand", ["All"] + sorted(d0["brand"].dropna().unique()))
if brand != "All":
    d0 = d0[d0["brand"] == brand]

destination = st.sidebar.selectbox("Destination", ["All"] + sorted(d0["destination"].dropna().unique()))
if destination != "All":
    d0 = d0[d0["destination"] == destination]

# Apply stored overrides first
d0["row_id"] = d0["month"].astype(str) + "||" + d0["brand"].astype(str) + "||" + d0["destination"].astype(str)
d0 = apply_overrides(d0, st.session_state.overrides)

# compute metrics AFTER overrides
d = compute_metrics(d0)

# -----------------------------
# Export filtered CSV
# -----------------------------
st.download_button(
    "Download filtered data (CSV)",
    d[FINAL_COL_ORDER].to_csv(index=False).encode("utf-8"),
    file_name=f"filtered_{month}_{brand}_{destination}.csv".replace(" ", "_"),
    mime="text/csv",
)

# -----------------------------
# Editable table (ONE click apply; ONE set of buttons)
# -----------------------------
st.markdown('<div class="card">', unsafe_allow_html=True)
st.markdown("### Editable Table (Impressions + Converted Leads)")
st.markdown(
    '<div class="muted">Edit <b>Impressions</b> and <b>Converted Leads</b>, then click <b>Apply changes</b>. '
    '<b>CPL</b> and <b>Conversion Rate</b> update automatically.</div>',
    unsafe_allow_html=True
)

with st.form("edit_form", clear_on_submit=False):
    table = d[FINAL_COL_ORDER].copy()
    table["conversion_rate"] = (table["conversion_rate"] * 100).round(2)  # show %

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
# Optional: raw table
# -----------------------------
with st.expander("Show final filtered dataset"):
    out = d[FINAL_COL_ORDER].copy()
    out["conversion_rate_%"] = (out["conversion_rate"] * 100).round(2)
    st.dataframe(out, use_container_width=True, hide_index=True)
