import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows

st.set_page_config(page_title="Backfill Report Tool", layout="wide")
st.title("📄 Backfill Report Generator")

# Initialize session state
if "page" not in st.session_state:
    st.session_state.page = "zqm"

# ----------------- PAGE 1: Upload ZQM -----------------
if st.session_state.page == "zqm":
    st.header("Step 1: Upload and Filter ZQM Job File")
    zqm_file = st.file_uploader("📁 Upload ZQM Job File", type=["xlsx"], key="zqm")

    if zqm_file:
        df = pd.read_excel(zqm_file)
        df.columns = df.columns.str.strip()

        # Apply filters
        filtered_df = df[
            (df["GR Qty"] == 0) &
            (df["Status"].str.contains("rel", case=False, na=False)) &
            (~df["Status"].str.contains("teco", case=False, na=False))
        ]

        st.success(f"✅ Filtered {len(filtered_df)} rows")
        st.dataframe(filtered_df)

        # Create downloadable Excel
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            filtered_df.to_excel(writer, index=False, sheet_name="Filtered")
        output.seek(0)

        st.download_button(
            label="📥 Download Filtered ZQM as Excel (paste into /n/scwm/mon)",
            data=output,
            file_name="filtered_zqm.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        if st.button("➡️ Proceed to Upload PMR File"):
            st.session_state.page = "pmr"
            st.rerun()

# ----------------- PAGE 2: Upload PMR -----------------
elif st.session_state.page == "pmr":
    st.header("Step 2: Upload PMR File")

    pmr_file = st.file_uploader("📁 Upload PMR File", type=["xlsx"], key="pmr")

    if pmr_file:
        pmr_df = pd.read_excel(pmr_file)
        st.success("✅ PMR file uploaded successfully!")
        st.dataframe(pmr_df)

        # ✅ Create first pivot table
        pivot1 = pd.pivot_table(
            pmr_df,
            index="Manufacturing Order",
            columns="Staging Status",
            values="Product",
            aggfunc="count",
            fill_value=0
        )
        pivot1["Grand Total"] = pivot1.sum(axis=1)
        pivot1.columns.name = None
        pivot1.reset_index(inplace=True)

        # ✅ Create second pivot table
        pivot2 = pd.pivot_table(
            pmr_df,
            index="Manufacturing Order",
            columns="Goods Issue Status",
            values="Product",
            aggfunc="count",
            fill_value=0
        )
        pivot2["Grand Total"] = pivot2.sum(axis=1)
        pivot2.columns.name = None
        pivot2.reset_index(inplace=True)

        # ✅ Filter pivot tables to include only rows with non-zero 'Completed' and 'Not Started'
        def has_both_statuses(df, cols):
            lower_cols = [c.lower() for c in cols]
            completed_col = next((c for c in df.columns if c.lower() == "completed"), None)
            not_started_col = next((c for c in df.columns if c.lower() == "not started"), None)
            if completed_col and not_started_col:
                return df[(df[completed_col] > 0) & (df[not_started_col] > 0)]
            else:
                return df.head(0)  # return empty DataFrame if columns don't exist

        pivot1 = has_both_statuses(pivot1, pivot1.columns)
        pivot2 = has_both_statuses(pivot2, pivot2.columns)

        # ✅ Write both sheets to Excel, with separate pivot tables on same sheet
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            pmr_df.to_excel(writer, index=False, sheet_name="Original PMR")
            pivot1.to_excel(writer, index=False, sheet_name="Pivot Summary")
            workbook = writer.book
            sheet = writer.sheets["Pivot Summary"]

            # Find the next empty column to place second pivot (leave a few columns of space)
            start_col = pivot1.shape[1] + 3

            for r_idx, row in enumerate(dataframe_to_rows(pivot2, index=False, header=True), 1):
                for c_idx, value in enumerate(row, start_col):
                    sheet.cell(row=r_idx, column=c_idx, value=value)

        output.seek(0)

        st.download_button(
            label="📥 Download PMR Output with Pivot",
            data=output,
            file_name="processed_pmr_with_pivot.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        if st.button("➡️ Proceed to Upload SOH File"):
            st.session_state.page = "soh"
            st.rerun()

# ----------------- PAGE 3: Upload SOH -----------------
elif st.session_state.page == "soh":
    st.header("Step 3: Upload SOH File")

    soh_file = st.file_uploader("📁 Upload SOH File", type=["xlsx"], key="soh")

    if soh_file:
        soh_df = pd.read_excel(soh_file)
        st.success("✅ SOH file uploaded successfully!")
        st.dataframe(soh_df)

        # Further merging/logic would go here

    if st.button("🔁 Restart Entire Process"):
        st.session_state.page = "zqm"
        st.rerun()
