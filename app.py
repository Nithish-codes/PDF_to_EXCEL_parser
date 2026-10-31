import streamlit as st
import pandas as pd
import tabula
import tempfile
import os

st.set_page_config(page_title="PDF to Excel Converter", page_icon="📄", layout="centered")

st.title("📊 PDF Statement to Excel Converter")
st.write("Upload your multi-page PDF bank statement below to convert and consolidate it into a clean Excel file.")

uploaded_file = st.file_uploader("Choose a PDF file", type=["pdf"])

if uploaded_file is not None:
    st.info(f"File uploaded: **{uploaded_file.name}** ({uploaded_file.size / 1024 / 1024:.2f} MB)")
    
    if st.button("Convert to Excel", type="primary"):
        with st.spinner("Processing PDF... For large files, this can take a minute."):
            with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_file:
                tmp_file.write(uploaded_file.getvalue())
                tmp_path = tmp_file.name

            output_path = tmp_path.replace(".pdf", ".xlsx")

            try:
                df_list = tabula.read_pdf(
                    tmp_path, 
                    pages='all', 
                    stream=True, 
                    multiple_tables=True,
                    java_options="-Xmx4g"
                )

                if not df_list:
                    st.error("No tables found in this PDF. It might be image-based (scanned).")
                else:
                    full_df = pd.concat(df_list, ignore_index=True)
                    full_df = full_df[full_df.iloc[:, 0] != "Date"]
                    full_df.dropna(how='all', inplace=True)

                    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                        full_df.to_excel(writer, sheet_name='Transactions', index=False)

                    st.success("Conversion successful!")

                    with open(output_path, "rb") as f:
                        st.download_button(
                            label="📥 Download Final_Consolidated_Statement.xlsx",
                            data=f,
                            file_name="Final_Consolidated_Statement.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )

            except Exception as e:
                st.error(f"An error occurred during processing: {e}")
                
            finally:
                if os.path.exists(tmp_path):
                    os.remove(tmp_path)
                if os.path.exists(output_path):
                    os.remove(output_path)