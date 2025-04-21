import streamlit as st
import pandas as pd
import requests
import io
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import numbers

# ---------------------------------------------
# Streamlit UI Enhancements
# ---------------------------------------------
st.markdown("""
    <style>
        .main {
            background-image: url("https://images.unsplash.com/photo-1532619187608-e5375cab36c9?ixlib=rb-4.0.3&auto=format&fit=crop&w=1950&q=80");
            background-size: cover;
            background-attachment: fixed;
        }

        .designer {
            text-align: center;
            font-size: 26px;
            color: #ffffff;
            font-weight: bold;
            margin-top: 20px;
            font-family: 'Arial', sans-serif;
        }

        .center-title {
            text-align: center;
            color: #ffffff;
            background-color: rgba(0, 0, 0, 0.6);
            padding: 1rem;
            border-radius: 10px;
            font-family: 'Arial', sans-serif;
            margin-top: 10px;
        }

        .dedication {
            text-align: center;
            font-size: 18px;
            margin-top: -10px;
            color: #dddddd;
        }
    </style>

    <div class="designer">👨‍💻 Designed by Engr. Ozair Khan</div>

    <div class="center-title">
        <h1>🔍 PESCO Bill Extractor Tool</h1>
        <p class="dedication">🎓 Dedicated to Engr. Bilal Shalman</p>
    </div>
""", unsafe_allow_html=True)

uploaded_file = st.file_uploader("📤 Upload your Excel file", type=["xlsx", "xls"])

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file, dtype=str)  # Read all as string to preserve formatting
        df = df.where(pd.notnull(df), "")  # Replace NaNs

        st.success("✅ File uploaded successfully.")
        st.write("📄 Preview of Uploaded Data:")
        st.dataframe(df)

        selected_col = st.selectbox("1️⃣ Select the column containing Account Numbers:", df.columns)

        target_col = st.selectbox("2️⃣ Select the column where Customer ID should be saved:",
                                  df.columns.tolist() + ["➕ Create new column..."])

        if target_col == "➕ Create new column...":
            new_col_name = st.text_input("Enter name for new column:")
            if new_col_name:
                if new_col_name not in df.columns:
                    df[new_col_name] = ""
                    target_col = new_col_name
                else:
                    st.warning("⚠️ Column already exists. Please choose another name.")
                    st.stop()

        if st.checkbox("⚠️ I understand this will modify the selected column with extracted data. Proceed?"):
            if st.button("🚀 Start Extracting Customer IDs"):
                with st.spinner("🔄 Extracting customer IDs via backend API..."):
                    customer_ids = []

                    for i, acc in enumerate(df[selected_col], start=1):
                        st.info(f"🔁 Processing {i} of {len(df)}...")  # Sequential status update
                        try:
                            acc_str = str(int(float(acc))).zfill(14)
                        except:
                            customer_ids.append("")
                            continue

                        if len(acc_str) != 14:
                            customer_ids.append("")
                            continue

                        # ✅ USE YOUR DEPLOYED API URL HERE
                        response = requests.post(
                            "https://backend-pescoapp-production.up.railway.app/get_customer_id",
                            json={"account_number": acc_str}
                        )

                        if response.status_code == 200:
                            cid = response.json().get("customer_id", "")
                            customer_ids.append(cid)
                        else:
                            customer_ids.append("")

                    df[target_col] = customer_ids

                st.success("✅ Extraction completed successfully.")
                st.write("🔎 Final Updated Data:")
                st.dataframe(df)

                @st.cache_data
                def to_excel_with_text_format(df, account_col):
                    wb = Workbook()
                    ws = wb.active
                    ws.title = "Data"

                    for r in dataframe_to_rows(df, index=False, header=True):
                        ws.append(r)

                    # Set Account Number column format to text
                    col_index = df.columns.get_loc(account_col) + 1  # openpyxl is 1-indexed
                    for cell in ws.iter_cols(min_col=col_index, max_col=col_index, min_row=2):
                        for c in cell:
                            c.number_format = '@'  # Set format to text

                    output = io.BytesIO()
                    wb.save(output)
                    return output.getvalue()

                excel_data = to_excel_with_text_format(df, selected_col)

                st.download_button(
                    label="📥 Download Updated Excel",
                    data=excel_data,
                    file_name="updated_data.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

    except Exception as e:
        st.error(f"❌ Error reading file: {e}")


