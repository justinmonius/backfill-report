import streamlit as st
import pandas as pd
from io import BytesIO

st.title("Backfill Report")

# Upload ZQM file
zqm_file = st.file_uploader("Upload ZQM Job File", type=["xlsx"])

if zqm_file:
    # Read Excel
    df = pd.read_excel(zqm_file)

    # Normalize column names
    df.columns = df.columns.str.strip()

    # Apply filters:
    # GR Qty == 0
    # Status contains "rel" (case-insensitive)
    # Status does NOT contain "teco"
    filtered_df = df[
        (df["GR Qty"] == 0) &
        (df["Status"].str.contains("rel", case=False, na=False)) &
        (~df["Status"].str.contains("teco", case=False, na=False))
    ]

    st.success(f"Filtered rows: {len(filtered_df)} found")

    # Show result
    st.dataframe(filtered_df)

    # Prepare Excel for download
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
else:
    st.info("Please upload a ZQM Job Excel file to begin.")
