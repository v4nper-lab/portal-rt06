import streamlit as st
import pandas as pd
import os

st.set_page_config(
    page_title="Portal Resmi RT 06 RW 14 Griya Permata Raya",
    layout="wide",
    page_icon="🏠"
)

st.title("🏠 Portal Resmi RT 06 / RW 14")
st.markdown("### Griya Permata Raya - Desa Nanjung Mekar")
st.write("---")

# Load data khusus RT 06
if os.path.exists("data_warga_rt06.xlsx"):
    df = pd.read_excel("data_warga_rt06.xlsx")
else:
    st.warning("⚠️ File 'data_warga_rt06.xlsx' belum di-upload ke GitHub.")
    df = pd.DataFrame()

if not df.empty:
    st.metric("👥 Total Warga RT 06 Terdaftar", f"{len(df)} Jiwa")
    st.subheader("📋 Data Warga RT 06")
    st.dataframe(df, use_container_width=True, hide_index=True)
else:
    st.info("Silakan upload file Excel data warga RT 06 ke repository ini.")
