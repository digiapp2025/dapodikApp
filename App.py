import streamlit as st


st.set_page_config(
    page_title="Monitoring Progres SYNC DAPODIK",
    layout="wide",
    initial_sidebar_state="expanded"    
)

pages = {
    "": [
        st.Page("pages/dashboard.py", title="Dashboard"),
    ],
    "Master": [
        st.Page("pages/master_SP.py", title="Master SP"),
        st.Page("pages/master_PD.py", title="Master PD"),
        st.Page("pages/master_GTK.py", title="Master GTK"),
    ],
    "Rekap": [
        st.Page("pages/rekap_Progres.py", title="Rekap Progres"),
        st.Page("pages/rekap_GTK.py", title="Rekap GTK"),
    ],
    "Pivot": [
        st.Page("pages/pivot_Progres.py", title="Pivot Progres"),
    ],
    "Tools": [
        st.Page("pages/merger_Excel.py", title="Merger Excel"),
        st.Page("pages/export_Data.py", title="Export Data"),
        st.Page("pages/import_Data.py", title="Import Data"),
    ],
}

pg = st.navigation(pages)
pg.run()