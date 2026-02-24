
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

st.title("Rekap GTK")

data = pd.DataFrame({
    "Kecamatan": ["A", "B", "C"],
    "Total GTK": [300, 250, 280],
    "ASN": [200, 150, 180],
    "Non ASN": [100, 100, 100]
})

col1, col2 = st.columns(2)

col1.metric("Total GTK", data["Total GTK"].sum())
col2.metric("Total ASN", data["ASN"].sum())

st.markdown("---")
st.dataframe(data, use_container_width=True)
