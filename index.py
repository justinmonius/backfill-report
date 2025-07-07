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
zqm_raw_df = None
if st.session_state.page == "zqm":
    st.header("Step 1: Upload and Filter ZQM Job File")
    zqm_file = st.file_uploader("📁 Upload ZQM Job File", type=["xlsx"], key="zqm")

    if zqm_file:
        df = pd.read_excel(zqm_file)
        df.columns = df.columns.str.strip()
        zqm_raw_df = df.copy()

        # Store ZQM file for reuse
        st.session_state.zqm_df = df

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

        # Drop specified columns from PMR
        columns_to_drop = [
            "Blocked (Overall)", "Stock Type", "Unit of Measure", "Document Number",
            "Operation or Activity", "Stor. Bin of Goods Mvt Posting", "Party Entitled to Dispose",
            "Production Supply Area", "Reservation Number", "Staging Method", "Requirement Start Date"
        ]
        pmr_df = pmr_df.drop(columns=[col for col in columns_to_drop if col in pmr_df.columns])

        # Add ZQM basic start/finish dates
        if "zqm_df" in st.session_state:
            zqm_df = st.session_state.zqm_df.copy()
            zqm_df.columns = zqm_df.columns.str.strip().str.lower()
            order_col = next((col for col in zqm_df.columns if "order" in col), None)
            start_col = next((col for col in zqm_df.columns if "start" in col), None)
            finish_col = next((col for col in zqm_df.columns if "finish" in col), None)
            if order_col and start_col and finish_col:
                zqm_df[start_col] = pd.to_datetime(zqm_df[start_col]).dt.strftime("%m/%d/%Y")
                zqm_df[finish_col] = pd.to_datetime(zqm_df[finish_col]).dt.strftime("%m/%d/%Y")
                zqm_subset = zqm_df[[order_col, start_col, finish_col]].drop_duplicates()
                zqm_subset.columns = ["Manufacturing Order", "Basic start date", "Basic finish date"]
                pmr_df = pmr_df.merge(zqm_subset, on="Manufacturing Order", how="left")

        st.session_state.pmr_df = pmr_df

        st.success("✅ PMR file uploaded successfully! Now upload SOH file.")
        st.dataframe(pmr_df)

        # Allow user to download raw cleaned PMR file
        pmr_output = BytesIO()
        with pd.ExcelWriter(pmr_output, engine='openpyxl') as writer:
            pmr_df.to_excel(writer, index=False, sheet_name="PMR_Cleaned")
        pmr_output.seek(0)

        st.download_button(
            label="📥 Download Cleaned PMR File",
            data=pmr_output,
            file_name="cleaned_pmr.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        if st.button("➡️ Proceed to Upload SOH File"):
            st.session_state.page = "soh"
            st.rerun()

# ----------------- PAGE 3: Upload SOH and Merge -----------------
elif st.session_state.page == "soh":
    st.header("Step 3: Upload SOH File")

    soh_file = st.file_uploader("📁 Upload SOH File", type=["xlsx"], key="soh")

    if soh_file and "pmr_df" in st.session_state:
        pmr_df = st.session_state.pmr_df.copy()
        soh_df = pd.read_excel(soh_file)

        # Filter SOH file
        soh_df.columns = soh_df.columns.str.strip()
        valid_stock_types = ["F1", "F2", "F3", "F4", "Q3", "Q4"]
        valid_storage_types = [
            "900", "900K", "905", "909", "910", "912", "914", "919", "920",
            "927", "934", "940", "986", "987", "990", "990P", "991", "FGI",
            "GRB1", "PRO", "PTWY", "RI", "RMA"
        ]
        soh_df_filtered = soh_df[
            soh_df["Stock Type"].isin(valid_stock_types) &
            soh_df["Storage Type"].astype(str).isin(valid_storage_types)
        ]

        # Create SOH pivot table
        soh_pivot = pd.pivot_table(
            soh_df_filtered,
            index="Product",
            columns="Owner",
            values="Quantity",
            aggfunc="sum",
            fill_value=0
        ).reset_index()

        # Lookup to add 9191 and 9192 values to PMR
        product_sums = soh_pivot.set_index("Product")[[col for col in soh_pivot.columns if col in ["MR9191", "MR9192"]]]
        pmr_df["9191"] = pmr_df["Product"].map(product_sums["MR9191"] if "MR9191" in product_sums else 0).fillna(0)
        pmr_df["9192"] = pmr_df["Product"].map(product_sums["MR9192"] if "MR9192" in product_sums else 0).fillna(0)

        # Create pivot tables
        pivot1 = pd.pivot_table(
            pmr_df,
            index="Manufacturing Order",
            columns="Staging Status",
            values="Product",
            aggfunc="count",
            fill_value=0
        )
        pivot1.columns.name = None
        pivot1.reset_index(inplace=True)

        pivot2 = pd.pivot_table(
            pmr_df,
            index="Manufacturing Order",
            columns="Goods Issue Status",
            values="Product",
            aggfunc="count",
            fill_value=0
        )
        pivot2.columns.name = None
        pivot2.reset_index(inplace=True)

        combined_df = pd.merge(pivot1, pivot2, on="Manufacturing Order", how="outer", suffixes=('_Staging', '_GI')).fillna(0)

        def determine_status(row):
            c_staging = row.get("Completed_Staging", 0)
            ns_staging = row.get("Not Started_Staging", 0)
            nr_staging = row.get("Not Relevant_Staging", 0)
            c_gi = row.get("Completed_GI", 0)
            ns_gi = row.get("Not Started_GI", 0)

            if c_staging == 0 and c_gi == 0:
                return "Not Pulled"
            elif (c_staging > 0 or c_gi > 0) and ns_staging == 0 and ns_gi == 0:
                return "Completed"
            elif (c_staging > 0 or c_gi > 0) and (ns_staging > 0 or ns_gi > 0):
                return "Partially Completed"
            else:
                return "Unknown"

        status_map = combined_df.set_index("Manufacturing Order").apply(determine_status, axis=1).to_dict()
        pmr_df.insert(0, "Hit", pmr_df["Manufacturing Order"].apply(lambda x: status_map.get(x, None)))

        st.success("✅ SOH processed and merged with PMR successfully!")
        st.dataframe(pmr_df)

        # Save to Excel
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            pmr_df.to_excel(writer, index=False, sheet_name="MASTER")
            soh_df_filtered.to_excel(writer, index=False, sheet_name="SOH Raw")
            soh_pivot.to_excel(writer, index=False, sheet_name="SOH Pivot")
            pivot1.to_excel(writer, index=False, sheet_name="Pivot Summary")
            combined_df.to_excel(writer, index=False, sheet_name="Combined Pivot")
            writer.book["MASTER"].sheet_properties.tabColor = "00FF00"
        output.seek(0)

        st.download_button(
            label="📥 Download Updated PMR File",
            data=output,
            file_name="updated_pmr_with_soh.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        if st.button("🔁 Restart Entire Process"):
            st.session_state.page = "zqm"
            st.rerun()
