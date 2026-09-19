import streamlit as st
import pandas as pd
import plotly.express as px
import os

st.set_page_config(
    page_title="Portal Resmi RT 06 RW 14 Griya Permata Raya",
    layout="wide",
    page_icon="🏠"
)

st.title("🏠 Portal Resmi RT 06 / RW 14")
st.markdown("### Griya Permata Raya - Desa Nanjung Mekar")
st.write("---")

# Load data khusus RT 06 dengan penyesuaian baris header Excel
@st.cache_data
def load_data_rt06():
    file_excel = "data_warga_rt06.xlsx"
    if not os.path.exists(file_excel):
        return pd.DataFrame()
    
    try:
        # Mencoba membaca file, kita lewati baris atas yang kosong/judul laporan (header=3 atau sesuaikan)
        # Berdasarkan gambar, header kolom yang benar ada di sekitar baris ke-4 (index 3 atau 4)
        df_raw = pd.read_excel(file_excel)
        
        # Cari baris yang berisi tulisan 'No.' atau 'Nama' untuk dijadikan header otomatis
        header_row = 0
        for idx, row in df_raw.iterrows():
            row_str = row.astype(str).values
            if any("nama" in val.lower() for val in row_str):
                header_row = idx
                break
        
        # Muat ulang dengan header yang benar
        df = pd.read_excel(file_excel, header=header_row)
        df.columns = df.columns.str.strip().str.upper()
        
        # Bersihkan baris kosong
        df = df.dropna(how="all")
        return df
    except Exception as e:
        st.error(f"Gagal memuat data: {e}")
        return pd.DataFrame()

df = load_data_rt06()

if not df.empty:
    st.metric("👥 Total Warga RT 06 Terdaftar", f"{len(df)} Jiwa")
    
    # Navigasi Tab
    tab_data, tab_grafik = st.tabs(["📋 Data Warga", "📈 Grafik Demografi Interaktif"])
    
    with tab_data:
        st.subheader("📋 Daftar Warga RT 06")
        st.dataframe(df, use_container_width=True, hide_index=True)
        
    with tab_grafik:
        st.subheader("📊 Statistik & Grafik Demografi Warga RT 06")
        
        # Deteksi kolom secara otomatis
        col_jk = next((c for c in df.columns if "JK" in c or "KELAMIN" in c or "GENDER" in c), None)
        col_pend = next((c for c in df.columns if "PENDIDIKAN" in c), None)
        col_pek = next((c for c in df.columns if "PEKERJAAN" in c), None)
        col_usia = next((c for c in df.columns if "USIA" in c or "UMUR" in c), None)
        col_status = next((c for c in df.columns if "STATUS" in c and "KAWIN" in c) or (c for c in df.columns if c == "STATUS"), None)
        col_domisili = next((c for c in df.columns if "DOMISILI" in c or "PENDUDUK" in c), None)
        col_rumah = next((c for c in df.columns if "RUMAH" in c or "HUNIAN" in c), None)
        
        c1, c2 = st.columns(2)
        
        # 1. Grafik Jenis Kelamin
        with c1:
            if col_jk:
                st.markdown("#### 👥 Berdasarkan Jenis Kelamin")
                df_jk = df[col_jk].value_counts().reset_index()
                df_jk.columns = ["Jenis Kelamin", "Jumlah"]
                fig_jk = px.pie(df_jk, names="Jenis Kelamin", values="Jumlah", hole=0.4, color_discrete_sequence=px.colors.qualitative.Set2)
                st.plotly_chart(fig_jk, use_container_width=True)
            else:
                st.info("Kolom Jenis Kelamin tidak terdeteksi.")
                
        # 2. Grafik Status Pernikahan
        with c2:
            if col_status:
                st.markdown("#### 💍 Berdasarkan Status Pernikahan")
                df_st = df[col_status].value_counts().reset_index()
                df_st.columns = ["Status", "Jumlah"]
                fig_st = px.bar(df_st, x="Status", y="Jumlah", text="Jumlah", color="Status", color_discrete_sequence=px.colors.qualitative.Pastel)
                st.plotly_chart(fig_st, use_container_width=True)
            else:
                st.info("Kolom Status Pernikahan tidak terdeteksi.")
                
        c3, c4 = st.columns(2)
        
        # 3. Grafik Pendidikan
        with c3:
            if col_pend:
                st.markdown("#### 🎓 Berdasarkan Pendidikan")
                df_pd = df[col_pend].value_counts().reset_index()
                df_pd.columns = ["Pendidikan", "Jumlah"]
                fig_pd = px.bar(df_pd, x="Pendidikan", y="Jumlah", text="Jumlah", color="Pendidikan", color_discrete_sequence=px.colors.qualitative.Bold)
                fig_pd.update_layout(xaxis=dict(tickangle=-30))
                st.plotly_chart(fig_pd, use_container_width=True)
            else:
                st.info("Kolom Pendidikan tidak terdeteksi.")
                
        # 4. Grafik Pekerjaan
        with c4:
            if col_pek:
                st.markdown("#### 💼 Berdasarkan Pekerjaan")
                df_pk = df[col_pek].value_counts().reset_index()
                df_pk.columns = ["Pekerjaan", "Jumlah"]
                fig_pk = px.bar(df_pk, x="Pekerjaan", y="Jumlah", text="Jumlah", color="Pekerjaan", color_discrete_sequence=px.colors.qualitative.Vivid)
                fig_pk.update_layout(xaxis=dict(tickangle=-30))
                st.plotly_chart(fig_pk, use_container_width=True)
            else:
                st.info("Kolom Pekerjaan tidak terdeteksi.")

        c5, c6 = st.columns(2)
        
        # 5. Grafik Kategori Usia
        with c5:
            if col_usia:
                st.markdown("#### 👶 Berdasarkan Kategori Usia")
                # Mengelompokkan usia secara otomatis
                def kategorikan_usia(u):
                    try:
                        u = int(u)
                        if u <= 5: return "Balita (0-5)"
                        elif u <= 12: return "Anak-anak (6-12)"
                        elif u <= 25: return "Remaja (13-25)"
                        elif u <= 50: return "Dewasa (26-50)"
                        else: return "Lansia (>50)"
                    except:
                        return "Tidak Diketahui"
                
                df_u = df.copy()
                df_u["KATEGORI_USIA"] = df_u[col_usia].apply(kategorikan_usia)
                df_usia_count = df_u["KATEGORI_USIA"].value_counts().reset_index()
                df_usia_count.columns = ["Kategori Usia", "Jumlah"]
                
                fig_usia = px.pie(df_usia_count, names="Kategori Usia", values="Jumlah", hole=0.4, color_discrete_sequence=px.colors.qualitative.Safe)
                st.plotly_chart(fig_usia, use_container_width=True)
            else:
                st.info("Kolom Usia tidak terdeteksi.")
                
        # 6 & 7. Domisili & Status Rumah (jika ada di kolom Excel)
        with c6:
            if col_domisili or col_rumah:
                if col_domisili:
                    st.markdown("#### 🏡 Status Domisili & Rumah")
                    df_dom = df[col_domisili].value_counts().reset_index()
                    df_dom.columns = ["Domisili", "Jumlah"]
                    fig_dom = px.bar(df_dom, x="Domisili", y="Jumlah", text="Jumlah", color_discrete_sequence=["#2563eb"])
                    st.plotly_chart(fig_dom, use_container_width=True)
            else:
                st.info("Kolom Status Domisili/Rumah belum ada di file Excel Anda (bisa ditambahkan nanti jika diperlukan).")
                
else:
    st.info("Silakan pastikan file Excel data warga RT 06 sudah di-upload dengan benar.")
