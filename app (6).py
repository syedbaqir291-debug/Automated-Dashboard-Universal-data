import streamlit as st
import pandas as pd

from dashboard_builder import (
    load_dataframe, clean_dataframe, profile_columns, build_workbook,
    MAX_DIMENSIONS,
)

st.set_page_config(page_title="Auto Pivot Dashboard Generator", layout="wide")

st.title("📊 Auto Pivot Dashboard Generator")
st.caption(
    "Upload any CSV / Excel file. The tool builds a workbook with native "
    "PivotTables (all sharing one PivotCache), KPI cards driven directly "
    "from those pivots, and matching charts — ready for you to drop "
    "Slicers onto in Excel for full cross-filtering."
)

uploaded = st.file_uploader("Upload data file", type=["csv", "xlsx", "xls"])

if uploaded is not None:
    try:
        df = load_dataframe(uploaded, uploaded.name)
    except Exception as e:
        st.error(f"Couldn't read file: {e}")
        st.stop()

    df = clean_dataframe(df)
    st.success(f"Loaded **{len(df):,}** rows × **{len(df.columns)}** columns")
    with st.expander("Preview data", expanded=False):
        st.dataframe(df.head(20), use_container_width=True)

    profile = profile_columns(df)

    if not profile['dimension_candidates']:
        st.error("Couldn't find any suitable breakdown columns in this file.")
        st.stop()

    st.subheader("1. Breakdown columns (dimensions)")
    st.caption(
        "These become your PivotTables, charts, and slicer fields. "
        "We've pre-selected good candidates — add or remove as needed "
        f"(max {MAX_DIMENSIONS})."
    )
    dims = st.multiselect(
        "Dimension columns",
        options=profile['dimension_candidates'],
        default=profile['suggested_dimensions'],
        max_selections=MAX_DIMENSIONS,
    )

    st.subheader("2. Measure (dependent variable)")
    measure_mode_label = st.radio(
        "What should each pivot/chart measure?",
        options=["Count of records", "Sum of a numeric column", "Average of a numeric column"],
        horizontal=True,
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

    st.subheader("3. Dashboard title")
    default_title = "Auto-Generated Dashboard"
    title = st.text_input("Title shown on the Dashboard sheet", value=default_title)
    source_label = st.text_input("Source label (optional, shown under the title)", value=uploaded.name)

    st.divider()

    if not dims:
        st.warning("Select at least one dimension column to continue.")
    else:
        if st.button("🚀 Generate Dashboard", type="primary"):
            with st.spinner("Building PivotTables, charts and KPI cards..."):
                try:
                    xlsx_bytes = build_workbook(
                        df, dims, measure_mode, measure_col,
                        title=title, source_label=source_label,
                    )
                except Exception as e:
                    st.error(f"Something went wrong while building the workbook: {e}")
                    st.stop()

            st.success("Dashboard ready!")
            st.download_button(
                "⬇️ Download Dashboard.xlsx",
                data=xlsx_bytes,
                file_name="Dashboard.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

            with st.expander("📌 One last step in Excel — adding Slicers", expanded=True):
                st.markdown(
                    """
The workbook already contains real, native PivotTables (one per
dimension you picked), all sharing a single PivotCache — that's what
makes cross-filtering possible.

To add the interactive filter buttons (Slicers):

1. Open the file in Excel, **unhide the `PivotData` sheet**
   (right-click any sheet tab → *Unhide*).
2. Click any cell inside one of the PivotTables.
3. **Insert → Slicer**, and tick the field(s) you want as filters
   (e.g. the dimension columns).
4. Right-click each slicer → **Report Connections** → tick *every*
   PivotTable listed.
5. Cut/paste the slicer(s) onto the **Dashboard** sheet, then
   re-hide `PivotData`.

Once connected, every chart *and* every KPI card (since they read
PivotTable cells directly) will update together when you click a
slicer.
                    """
                )
else:
    st.info("👆 Upload a CSV or Excel file to get started.")
