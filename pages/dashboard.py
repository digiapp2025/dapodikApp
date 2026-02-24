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

st.title("Dashboard Monitoring DAPODIK")


# =========================
# DATA SIMULASI (ganti dengan real data nanti)
# =========================
total_sekolah = 1250
sudah_sync = 980
bulan_lalu_sync = 955

persen_sync = round((sudah_sync / total_sekolah) * 100, 2)
delta_sync = sudah_sync - bulan_lalu_sync

# =========================
# KPI
# =========================
col1, col2, col3 = st.columns(3)

col1.metric(
    "Total Sekolah",
    f"{total_sekolah:,}"
)

col2.metric(
    "Sudah SYNC",
    f"{sudah_sync:,}",
    f"+{delta_sync:,}"
)

col3.metric(
    "Persentase SYNC",
    f"{persen_sync}%"
)

st.divider()

# =========================
# GRAFIK PROGRES
# =========================
st.subheader("📈 Grafik Progres SYNC")

progres = pd.DataFrame({
    "Bulan": ["Jan", "Feb", "Mar", "Apr"],
    "Sudah SYNC": [850, 900, 955, 980]
})

st.line_chart(
    progres.set_index("Bulan"),
    use_container_width=True
)

# =========================
# RANKING KECAMATAN
# =========================
st.subheader("🏆 Ranking Kecamatan")

ranking = pd.DataFrame({
    "Kecamatan": ["A", "B", "C"],
    "Total Sekolah": [120, 150, 100],
    "Sudah SYNC": [110, 132, 85]
})

ranking["% Sync"] = (
    ranking["Sudah SYNC"] / ranking["Total Sekolah"] * 100
).round(2)

ranking = ranking.sort_values("% Sync", ascending=False)

st.dataframe(
    ranking,
    use_container_width=True
)
