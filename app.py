import pandas as pd
import streamlit as st

from dashboard_builder import (
    load_dataframe, clean_dataframe, profile_columns, build_workbook,
    MAX_DIMENSIONS, BLANK, BRAND, using_macro_template,
)

NAVY = "#1F4E5C"
TEAL = "#2E8B8B"
GOLD = "#D4AC0D"
ACCENT = "#C0392B"
LIGHT = "#EAF2F2"

st.set_page_config(page_title="Auto Pivot Dashboard Generator", page_icon="📊", layout="wide")

# ----------------------------------------------------------------------
# Styling
# ----------------------------------------------------------------------
st.markdown(f"""
<style>
.stApp {{ background-color: #F7FAFA; }}

.omac-header {{
    background: linear-gradient(135deg, {NAVY} 0%, {TEAL} 100%);
    border-radius: 14px;
    padding: 1.4rem 1.8rem;
    color: white;
    display: flex;
    justify-content: space-between;
    align-items: center;
    margin-bottom: 1.2rem;
    box-shadow: 0 4px 14px rgba(31,78,92,0.25);
}}
.omac-header h1 {{ margin: 0; font-size: 1.7rem; }}
.omac-header p {{ margin: 0.15rem 0 0 0; opacity: 0.85; font-size: 0.95rem; }}
.omac-badge {{
    background: {GOLD};
    color: {NAVY};
    font-weight: 700;
    font-size: 0.8rem;
    padding: 0.35rem 0.9rem;
    border-radius: 999px;
    white-space: nowrap;
    box-shadow: 0 2px 6px rgba(0,0,0,0.15);
}}

.section-card {{
    background: white;
    border-radius: 12px;
    padding: 1.2rem 1.4rem;
    margin-bottom: 1rem;
    box-shadow: 0 2px 10px rgba(0,0,0,0.05);
    border: 1px solid #E5EEEE;
}}
.section-card h3 {{ margin-top: 0; color: {NAVY}; }}

div[data-testid="stFileUploaderDropzone"] {{
    background: {LIGHT};
    border: 2px dashed {TEAL};
    border-radius: 12px;
}}

.stButton button[kind="primary"] {{
    background: linear-gradient(135deg, {NAVY} 0%, {TEAL} 100%);
    border: none;
    border-radius: 8px;
    padding: 0.6rem 1.6rem;
    font-weight: 700;
    box-shadow: 0 3px 8px rgba(31,78,92,0.3);
}}

.omac-footer {{
    text-align: center;
    padding: 1rem;
    margin-top: 2rem;
    color: {NAVY};
    font-size: 0.85rem;
    border-top: 1px solid #E5EEEE;
}}
.omac-footer .pill {{
    display: inline-block;
    background: {GOLD};
    color: {NAVY};
    font-weight: 700;
    padding: 0.25rem 0.8rem;
    border-radius: 999px;
    margin-top: 0.3rem;
}}
</style>
""", unsafe_allow_html=True)

# ----------------------------------------------------------------------
# Header
# ----------------------------------------------------------------------
st.markdown(f"""
<div class="omac-header">
    <div>
        <h1>📊 Auto Pivot Dashboard Generator</h1>
        <p>Upload any data &middot; pick dimensions &amp; a measure &middot; get a polished
        workbook with KPI cards, charts and a clean summary sheet &mdash; ready to open,
        no Excel repair prompts.</p>
    </div>
    <div class="omac-badge">{BRAND}</div>
</div>
""", unsafe_allow_html=True)

# ----------------------------------------------------------------------
# 1. Upload
# ----------------------------------------------------------------------
st.markdown('<div class="section-card">', unsafe_allow_html=True)
st.markdown("### 1️⃣ Upload your data")
uploaded = st.file_uploader(
    "Drag & drop a CSV or Excel file here",
    type=["csv", "xlsx", "xls"],
    label_visibility="collapsed",
)
st.markdown('</div>', unsafe_allow_html=True)

if uploaded is None:
    st.info("👆 Upload a file to get started — try your complaint register, sales log, "
            "attendance sheet... whatever you've got. We'll figure out the rest.")
    st.markdown(f"""
    <div class="omac-footer">
        Built for healthcare quality &amp; analytics teams<br>
        <span class="pill">{BRAND}</span>
    </div>
    """, unsafe_allow_html=True)
    st.stop()

try:
    df = load_dataframe(uploaded, uploaded.name)
except Exception as e:
    st.error(f"Couldn't read this file: {e}")
    st.stop()

df = clean_dataframe(df)
profile = profile_columns(df)

st.markdown('<div class="section-card">', unsafe_allow_html=True)
m1, m2, m3, m4 = st.columns(4)
m1.metric("Rows", f"{len(df):,}")
m2.metric("Columns", f"{len(df.columns)}")
m3.metric("Dimension candidates", f"{len(profile['dimension_candidates'])}")
m4.metric("Numeric columns", f"{len(profile['numeric_candidates'])}")
with st.expander("Preview data"):
    st.dataframe(df.head(20), use_container_width=True)
st.markdown('</div>', unsafe_allow_html=True)

if not profile['dimension_candidates']:
    st.error("Couldn't find any suitable breakdown columns in this file — "
             "try a file with at least one categorical / low-cardinality column.")
    st.stop()

# ----------------------------------------------------------------------
# 2. Configuration
# ----------------------------------------------------------------------
st.markdown('<div class="section-card">', unsafe_allow_html=True)
st.markdown("### 2️⃣ Choose breakdown columns (dimensions)")
st.caption(
    f"These become your summary tables, charts and KPI cards. We've pre-selected good "
    f"candidates — add or remove as needed (max {MAX_DIMENSIONS})."
)
dims = st.multiselect(
    "Dimension columns",
    options=profile['dimension_candidates'],
    default=profile['suggested_dimensions'],
    max_selections=MAX_DIMENSIONS,
    label_visibility="collapsed",
)
st.markdown('</div>', unsafe_allow_html=True)

chart_types = {}
if dims:
    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.markdown("### 🍩 Chart style per dimension")
    st.caption(
        "By default, dimensions with 5 or fewer categories become donut charts "
        "and the rest become bar charts. Override any of these below — useful "
        "for things like 'Month' where you might prefer a bar/trend view even "
        "with few categories."
    )
    cols = st.columns(min(len(dims), 4))
    for i, d in enumerate(dims):
        with cols[i % len(cols)]:
            choice = st.selectbox(
                d,
                options=["Auto", "Donut", "Bar"],
                key=f"chart_type_{d}",
            )
            chart_types[d] = choice.lower()
    st.markdown('</div>', unsafe_allow_html=True)

st.markdown('<div class="section-card">', unsafe_allow_html=True)
st.markdown("### 3️⃣ Choose the measure (dependent variable)")
measure_mode_label = st.radio(
    "What should each summary table/chart measure?",
    options=["Count of records", "Sum of a numeric column", "Average of a numeric column"],
    horizontal=True,
    label_visibility="collapsed",
)

measure_col = None
if measure_mode_label != "Count of records":
    if not profile['numeric_candidates']:
        st.warning("No numeric columns detected — falling back to Count of records.")
        measure_mode_label = "Count of records"
    else:
        measure_col = st.selectbox("Numeric column", options=profile['numeric_candidates'])

measure_mode = {
    "Count of records": "count",
    "Sum of a numeric column": "sum",
    "Average of a numeric column": "average",
}[measure_mode_label]
st.markdown('</div>', unsafe_allow_html=True)

st.markdown('<div class="section-card">', unsafe_allow_html=True)
st.markdown("### 4️⃣ Title your dashboard")
col_a, col_b = st.columns(2)
title = col_a.text_input("Dashboard title", value="Auto-Generated Dashboard")
source_label = col_b.text_input("Source label (optional)", value=uploaded.name)
st.markdown('</div>', unsafe_allow_html=True)

# ----------------------------------------------------------------------
# 3. Generate + interactive preview
# ----------------------------------------------------------------------
if not dims:
    st.warning("Select at least one dimension column to continue.")
    st.stop()

generate = st.button("🚀 Generate Dashboard", type="primary", use_container_width=True)

if generate:
    with st.spinner("Building summary tables, charts and KPI cards..."):
        try:
            xlsx_bytes = build_workbook(
                df, dims, measure_mode, measure_col,
                title=title, source_label=source_label,
                chart_types=chart_types,
            )
        except Exception as e:
            st.error(f"Something went wrong while building the workbook: {e}")
            st.stop()

    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.markdown("### ✅ Output preview")

    # ---- live KPI preview (mirrors what the Excel KPI cards will show) ----
    if measure_mode == 'count':
        agg_label, agg_fmt = "Records", "{:,.0f}"
        def agg(d): return df[d].fillna(BLANK).astype(str).value_counts()
        total_val = len(df)
        total_label = "Total Records"
    elif measure_mode == 'sum':
        agg_label, agg_fmt = f"Sum of {measure_col}", "{:,.2f}"
        def agg(d): return df.assign(**{d: df[d].fillna(BLANK).astype(str)}).groupby(d)[measure_col].sum(min_count=1)
        total_val = df[measure_col].sum()
        total_label = f"Total {measure_col}"
    else:
        agg_label, agg_fmt = f"Average {measure_col}", "{:,.2f}"
        def agg(d): return df.assign(**{d: df[d].fillna(BLANK).astype(str)}).groupby(d)[measure_col].mean()
        total_val = df[measure_col].mean()
        total_label = f"Overall Average {measure_col}"

    d0 = dims[0]
    s0 = agg(d0).fillna(0).sort_values(ascending=False)

    k1, k2, k3, k4 = st.columns(4)
    k1.metric(total_label, agg_fmt.format(total_val))
    k2.metric(f"Top {d0}", str(s0.index[0]))
    k3.metric(f"{agg_label} ({d0} top)", agg_fmt.format(s0.iloc[0]))
    k4.metric(f"{d0} categories", f"{len(s0)}")

    st.caption("Charts below mirror the charts in the generated workbook "
               "(top categories shown):")

    chart_cols = st.columns(3)
    for i, d in enumerate(dims):
        s = agg(d).fillna(0).sort_values(ascending=False).head(10)
        with chart_cols[i % 3]:
            st.markdown(f"**{d}**")
            st.bar_chart(s)

    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.markdown("### ⬇️ Download")

    macro_enabled = using_macro_template()
    out_name = "Dashboard.xlsm" if macro_enabled else "Dashboard.xlsx"
    out_mime = (
        "application/vnd.ms-excel.sheet.macroEnabled.12"
        if macro_enabled else
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    st.download_button(
        f"Download {out_name}",
        data=xlsx_bytes,
        file_name=out_name,
        mime=out_mime,
        type="primary",
        use_container_width=True,
    )

    if macro_enabled:
        with st.expander("📌 About the downloaded workbook", expanded=True):
            st.markdown("""
The workbook has 4 sheets: **Dashboard** (KPI cards + charts), **Summary
Data** (per-category tables behind the KPIs), **Data** (cleaned source as
an Excel Table), and a hidden **Config** sheet.

**On first open, Excel will show a yellow security bar — click "Enable
Content".** This runs a built-in macro that automatically builds real
PivotTables, PivotCharts and **Slicers** (one per dimension, all
cross-connected) on a new "Live Pivots" area and on the Dashboard. Use the
slicers to filter everything at once — that's the genuinely interactive
view.

The macro re-runs (and rebuilds cleanly, with no duplicates) every time
the file is opened, so it always reflects the data in the **Data** sheet.
            """)
    else:
        with st.expander("📌 About the downloaded workbook", expanded=True):
            st.markdown("""
The workbook has 3 sheets: **Dashboard** (KPI cards + charts), **Summary
Data** (the per-category tables that drive every chart and KPI on the
Dashboard), and **Data** (cleaned source as an Excel Table).

This file opens cleanly in Excel with no "repair" prompts. If you'd like
to explore the raw data interactively yourself:

1. Go to the **Data** sheet and click anywhere inside the table.
2. **Insert → PivotTable** and choose where to place it.
3. Drag any field into Rows/Values to build your own pivot — Excel will
   create a fresh, valid PivotTable from this table.

For interactive, slicer-style filtering without leaving the browser, use
this app's own charts and KPI preview above — they update live as you
change your dimension/measure selections.

*Tip: set up `dashboard_macro_template.xlsm` (see
`ThisWorkbook_macro.vba`) to enable automatic, slicer-connected
PivotTables and PivotCharts in every downloaded dashboard.*
            """)
    st.markdown('</div>', unsafe_allow_html=True)

st.markdown(f"""
<div class="omac-footer">
    Built for healthcare quality &amp; analytics teams<br>
    <span class="pill">{BRAND}</span>
</div>
""", unsafe_allow_html=True)
