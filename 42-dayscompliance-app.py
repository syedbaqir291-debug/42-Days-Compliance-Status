import streamlit as st
import pandas as pd
from io import BytesIO

st.title("42 Days Compliance Checker - OMAC Developer by S M Baqir")

# -------- 1. Upload Excel File --------
uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])

if uploaded_file:
    xls = pd.ExcelFile(uploaded_file)
    sheets = xls.sheet_names

    # -------- 2. Select Multiple Sheets --------
    selected_sheets = st.multiselect(
        "Select Sheets to Process",
        sheets
    )

    if selected_sheets:
        # -------- 3. Compliance Criteria --------
        criteria_type = st.radio(
            "Select Compliance Criteria",
            ["Greater Than", "Less Than"]
        )
        
        days_value = st.number_input(
            "Enter number of days",
            min_value=0,
            step=1
        )

        # Process each selected sheet
        processed_data = {}

        for sheet in selected_sheets:
            df = pd.read_excel(uploaded_file, sheet_name=sheet)

            # -------- 4. Select Column to Apply Criteria --------
            st.write(f"### Column Selection for Sheet: **{sheet}**")
            column_choice = st.selectbox(
                f"Select column where Day values are present (Sheet: {sheet})",
                df.columns,
                key=sheet
            )

            # -------- 5. Apply Criteria & Create Status Column --------
            if criteria_type == "Greater Than":
                df["Status"] = df[column_choice].apply(
                    lambda x: "Met" if x > days_value else "Not Met"
                )
            else:
                df["Status"] = df[column_choice].apply(
                    lambda x: "Met" if x < days_value else "Not Met"
                )

            processed_data[sheet] = df

        # -------- 6. Download Updated Excel --------
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            for sheet, df in processed_data.items():
                df.to_excel(writer, sheet_name=sheet, index=False)

        st.success("Processing Completed!")

        st.download_button(
            label="Download Updated Excel File",
            data=output.getvalue(),
            file_name="Updated_Compliance_Check.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        # -------- 7. Show results --------
        for sheet, df in processed_data.items():
            st.write(f"### Preview: {sheet}")
            st.dataframe(df)
