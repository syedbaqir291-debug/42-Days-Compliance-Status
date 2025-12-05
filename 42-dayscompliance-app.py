import streamlit as st
import pandas as pd
from io import BytesIO
import base64

# -------------------- PREMIUM PAGE SETUP --------------------
st.set_page_config(page_title="Excel Compliance Checker - OMAC Developer by S M Baqir", layout="wide")

st.markdown(
    """
    <style>
        .main {background-color: #F9FAFB;}
        .title {font-size: 42px; font-weight: 700; color:#4F46E5; text-align:center;}
        .subtitle {font-size: 18px; color:#374151; text-align:center; margin-top:-10px;}
        .card {
            background: white; 
            padding: 25px; 
            border-radius: 18px; 
            box-shadow: 0px 4px 16px rgba(0,0,0,0.06);
            margin-bottom: 25px;
        }
        .stDownloadButton button {
            background-color:#4F46E5;
            color:white;
            border-radius:10px;
            padding:10px 20px;
            font-weight:600;
        }
    </style>
    """,
    unsafe_allow_html=True
)

st.markdown('<div class="title">Excel Compliance Checker</div>', unsafe_allow_html=True)
st.markdown('<div class="subtitle">Smart Rules • Multi-Sheet • Premium Interface</div>', unsafe_allow_html=True)
st.write("")

# -------------------- 1. UPLOAD FILE --------------------
with st.container():
    st.markdown('<div class="card">', unsafe_allow_html=True)
    uploaded_file = st.file_uploader("📤 Upload Excel Workbook (.xlsx)", type=["xlsx"])
    st.markdown("</div>", unsafe_allow_html=True)

if uploaded_file:
    xls = pd.ExcelFile(uploaded_file)
    sheet_options = xls.sheet_names

    # -------------------- 2. SELECT SHEETS --------------------
    st.markdown('<div class="card">', unsafe_allow_html=True)
    selected_sheets = st.multiselect(
        "📑 Select Sheets to Process",
        sheet_options
    )
    st.markdown("</div>", unsafe_allow_html=True)

    if selected_sheets:
        # -------------------- 3. COMPLIANCE CRITERIA --------------------
        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.markdown("### 📏 Compliance Criteria")

        criteria_type = st.radio(
            "Choose Criteria Type:",
            ["Greater Than", "Less Than"],
            horizontal=True
        )

        days_value = st.number_input(
            "Enter Number of Days:",
            min_value=0,
            step=1
        )

        st.markdown("</div>", unsafe_allow_html=True)

        processed_data = {}

        # -------------------- 4. PROCESS EACH SHEET --------------------
        for sheet in selected_sheets:

            st.markdown('<div class="card">', unsafe_allow_html=True)
            st.markdown(f"### 📝 Sheet: **{sheet}**")

            df = pd.read_excel(uploaded_file, sheet_name=sheet)

            # Column selector
            column_choice = st.selectbox(
                f"Select Column for Criteria (Sheet: {sheet})",
                df.columns,
                key=f"col_{sheet}"
            )

            # Blank value treatment
            blank_choice = st.selectbox(
                "How should blank values be treated?",
                ["Met", "Not Met", "Not Applicable"],
                key=f"blank_{sheet}"
            )

            # -------------------- APPLY LOGIC --------------------
            def evaluate_status(value):
                # Handle blanks
                if pd.isna(value) or value == "":
                    return blank_choice  

                # Handle non-numeric
                try:
                    value = float(value)
                except:
                    return "Not Applicable"

                # Compare
                if criteria_type == "Greater Than":
                    return "Met" if value > days_value else "Not Met"
                else:
                    return "Met" if value < days_value else "Not Met"

            df["Status"] = df[column_choice].apply(evaluate_status)
            processed_data[sheet] = df

            st.dataframe(df, use_container_width=True)
            st.markdown("</div>", unsafe_allow_html=True)

        # -------------------- 6. DOWNLOAD UPDATED EXCEL --------------------
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            for sheet_name, df in processed_data.items():
                df.to_excel(writer, sheet_name=sheet_name, index=False)

        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.success("✅ Processing Completed! Download your updated file below.")

        st.download_button(
            label="⬇️ Download Updated Excel",
            data=output.getvalue(),
            file_name="Updated_Compliance_Check.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        st.markdown("</div>", unsafe_allow_html=True)
