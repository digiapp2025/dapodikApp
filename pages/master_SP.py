
import streamlit as st
import pandas as pd

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

st.title("Master Satuan Pendidikan")

uploaded = st.file_uploader("Upload Master GTK", type=["xlsx", "csv"])

if uploaded:
    df = pd.read_excel(uploaded) if uploaded.name.endswith("xlsx") else pd.read_csv(uploaded)

    st.subheader("Preview Data")
    st.dataframe(df.head(), use_container_width=True)

    st.subheader("Statistik Ringkas")
    st.write(df.describe())
