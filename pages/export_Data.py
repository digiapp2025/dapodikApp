import streamlit as st
import pandas as pd
from io import BytesIO

st.markdown(
    """
    <style>
        .block-container {
            padding-top: 2.5rem;
        }
    </style>
    """,
    unsafe_allow_html=True
)

st.title("Export Data - File Excel")

# Upload multiple file
uploaded_files = st.file_uploader(
    "Upload semua file Excel",
    type=["xlsx", "xls"],
    accept_multiple_files=True
)

if uploaded_files:

    if st.button("🔄 Proses Merge"):

        df_list = []

        for file in uploaded_files:
            try:
                st.write(f"Membaca file: {file.name}")
                df = pd.read_excel(file, sheet_name="Sheet1")
                df["__source_file"] = file.name
                df_list.append(df)
            except Exception as e:
                st.error(f"Gagal membaca {file.name}: {e}")

        if df_list:

            combined_df = pd.concat(df_list, ignore_index=True)

            st.success(f"✅ Total file digabung: {len(df_list)}")
            st.write("Preview Data:")
            st.dataframe(combined_df)

            # Simpan ke memory buffer (bukan ke folder)
            output = BytesIO()
            combined_df.to_excel(output, index=False, engine='openpyxl')
            output.seek(0)

            st.download_button(
                label="📥 Download Hasil Merge",
                data=output,
                file_name="hasil_merge.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
