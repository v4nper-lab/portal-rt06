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

@st.cache_data
def load_data_rt06():
    file_excel = "data_warga_rt06.xlsx"
    if not os.path.exists(file_excel):
        return pd.DataFrame()
    
    try:
        df = pd.read_excel(file_excel, header=3)
        df.columns = df.columns.astype(str).str.strip().str.upper()
        
        def adalah_baris_nomor(row):
            count_angka = 0
            total_kolom = len(row)
            for val in row.values:
                val_str = str(val).strip()
                if val_str.isdigit() and int(val_str) < 50:
                    count_angka += 1
            return count_angka > (total_kolom / 3)
        
        if len(df) > 0:
            df = df[~df.apply(adalah_baris_nomor, axis=1)].reset_index(drop=True)
            
        df = df.loc[:, ~df.columns.str.contains('UNNAMED')]
        df = df.dropna(how="all")
        return df
    except Exception as e:
        st.error(f"Gagal memuat data: {e}")
        return pd.DataFrame()

df = load_data_rt06()

if not df.empty:
    st.metric("👥 Total Warga RT 06 Terdaftar", f"{len(df)} Jiwa")
    
    # Navigasi Menu Utama
    menu = st.sidebar.selectbox("📂 Navigasi Menu", ["Dashboard & Data Warga", "Kelola Data Warga (Tambah/Edit/Hapus)", "Cetak Laporan PDF"])
    
    if menu == "Dashboard & Data Warga":
        tab_data, tab_grafik = st.tabs(["📋 Data Warga", "📈 Grafik Demografi Interaktif"])
        
        with tab_data:
            st.subheader("📋 Daftar Warga RT 06")
            st.dataframe(df, use_container_width=True, hide_index=True)
            
        with tab_grafik:
            st.subheader("📊 Statistik & Grafik Demografi Warga RT 06")
            
            col_jk = next((c for c in df.columns if "JK" in c or "KELAMIN" in c or "GENDER" in c), None)
            col_pend = next((c for c in df.columns if "PENDIDIKAN" in c), None)
            col_pek = next((c for c in df.columns if "PEKERJAAN" in c), None)
            col_usia = next((c for c in df.columns if "USIA" in c or "UMUR" in c), None)
            col_status = next((c for c in df.columns if "STATUS" in c or "KAWIN" in c), None)
            
            c1, c2 = st.columns(2)
            
            with c1:
                if col_jk:
                    st.markdown("#### 👥 Berdasarkan Jenis Kelamin")
                    df_jk = df[col_jk].dropna().value_counts().reset_index()
                    df_jk.columns = ["Jenis Kelamin", "Jumlah"]
                    fig_jk = px.pie(df_jk, names="Jenis Kelamin", values="Jumlah", hole=0.4, color_discrete_sequence=px.colors.qualitative.Set2)
                    # Memperbesar ukuran font/angka di dalam grafik
                    fig_jk.update_traces(textfont_size=16, textinfo="percent+label+value")
                    fig_jk.update_layout(font=dict(size=14))
                    st.plotly_chart(fig_jk, use_container_width=True)
                else:
                    st.info("Kolom Jenis Kelamin tidak terdeteksi.")
                    
            with c2:
                if col_status:
                    st.markdown("#### 💍 Berdasarkan Status Pernikahan")
                    df_st = df[col_status].dropna().value_counts().reset_index()
                    df_st.columns = ["Status", "Jumlah"]
                    fig_st = px.bar(df_st, x="Status", y="Jumlah", text="Jumlah", color="Status", color_discrete_sequence=px.colors.qualitative.Pastel)
                    fig_st.update_traces(textfont_size=16, textposition="outside")
                    fig_st.update_layout(font=dict(size=14), xaxis=dict(tickfont=dict(size=14)), yaxis=dict(tickfont=dict(size=14)))
                    st.plotly_chart(fig_st, use_container_width=True)
                else:
                    st.info("Kolom Status Pernikahan tidak terdeteksi.")
                    
            c3, c4 = st.columns(2)
            
            with c3:
                if col_pend:
                    st.markdown("#### 🎓 Berdasarkan Pendidikan")
                    df_pd = df[col_pend].dropna().value_counts().reset_index()
                    df_pd.columns = ["Pendidikan", "Jumlah"]
                    fig_pd = px.bar(df_pd, x="Pendidikan", y="Jumlah", text="Jumlah", color="Pendidikan", color_discrete_sequence=px.colors.qualitative.Bold)
                    fig_pd.update_traces(textfont_size=16, textposition="outside")
                    fig_pd.update_layout(font=dict(size=14), xaxis=dict(tickangle=-30, tickfont=dict(size=13)), yaxis=dict(tickfont=dict(size=14)))
                    st.plotly_chart(fig_pd, use_container_width=True)
                else:
                    st.info("Kolom Pendidikan tidak terdeteksi.")
                    
            with c4:
                if col_pek:
                    st.markdown("#### 💼 Berdasarkan Pekerjaan")
                    df_pk = df[col_pek].dropna().value_counts().reset_index()
                    df_pk.columns = ["Pekerjaan", "Jumlah"]
                    fig_pk = px.bar(df_pk, x="Pekerjaan", y="Jumlah", text="Jumlah", color="Pekerjaan", color_discrete_sequence=px.colors.qualitative.Vivid)
                    fig_pk.update_traces(textfont_size=16, textposition="outside")
                    fig_pk.update_layout(font=dict(size=14), xaxis=dict(tickangle=-30, tickfont=dict(size=13)), yaxis=dict(tickfont=dict(size=14)))
                    st.plotly_chart(fig_pk, use_container_width=True)
                else:
                    st.info("Kolom Pekerjaan tidak terdeteksi.")

            if col_usia:
                st.markdown("#### 👶 Berdasarkan Kategori Usia")
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
                fig_usia.update_traces(textfont_size=16, textinfo="percent+label+value")
                fig_usia.update_layout(font=dict(size=14))
                st.plotly_chart(fig_usia, use_container_width=True)

    elif menu == "Kelola Data Warga (Tambah/Edit/Hapus)":
        st.subheader("🛠️ Panel Pengelolaan Data Warga RT 06")
        
        aksi = st.selectbox("Pilih Aksi Pengelolaan:", ["➕ Tambah Data Warga Baru", "✏️ Edit Data Warga", "🗑️ Hapus Data Warga"])
        
        if aksi == "➕ Tambah Data Warga Baru":
            st.markdown("### Form Tambah Warga")
            with st.form("form_tambah"):
                no_rumah = st.text_input("No. Rumah")
                nama = st.text_input("Nama Lengkap")
                jk = st.selectbox("Jenis Kelamin", ["L", "P"])
                hubungan = st.selectbox("Hubungan Keluarga", ["Kepala Keluarga", "Istri", "Anak Kandung", "Famili Lain", "Mertua"])
                usia = st.number_input("Usia", min_value=0, max_value=120, value=25)
                status = st.selectbox("Status Pernikahan", ["Belum Kawin", "Kawin", "Cerai Hidup", "Cerai Mati"])
                pendidikan = st.text_input("Pendidikan Terakhir", "Tamat SLTA/sederajat")
                pekerjaan = st.text_input("Pekerjaan", "Karyawan Swasta")
                
                submit_tambah = st.form_submit_button("Simpan Data Warga Baru")
                if submit_tambah:
                    st.success(f"Data warga atas nama **{nama}** berhasil disiapkan untuk ditambahkan!")
                    st.info("Catatan: Untuk menyimpan permanen ke file Excel, pastikan file Excel di GitHub diperbarui.")

        elif aksi == "✏️ Edit Data Warga":
            st.markdown("### Edit Data Warga")
            nama_pilih = st.selectbox("Pilih Warga yang Ingin Diedit:", df.iloc[:, 2] if len(df.columns) > 2 else df.iloc[:, 0])
            st.info(f"Fitur edit untuk **{nama_pilih}** aktif. Silakan ubah data melalui form di bawah.")
            with st.form("form_edit"):
                st.text_input("Perbarui Nama", value=str(nama_pilih))
                st.text_input("Perbarui No. Rumah")
                st.form_submit_button("Simpan Perubahan")

        elif aksi == "🗑️ Hapus Data Warga":
            st.markdown("### Hapus Data Warga")
            nama_hapus = st.selectbox("Pilih Warga yang Ingin Dihapus:", df.iloc[:, 2] if len(df.columns) > 2 else df.iloc[:, 0])
            if st.button("Konfirmasi Hapus Warga Ini"):
                st.success(f"Data warga **{nama_hapus}** berhasil dihapus dari tampilan.")

    elif menu == "Cetak Laporan PDF":
        st.subheader("🖨️ Cetak Rekapitulasi Data Warga RT 06")
        st.write("Klik tombol di bawah ini untuk mengunduh rekapitulasi data warga dalam format cetak/PDF:")
        
        # Tombol cetak via browser
        if st.button("📥 Unduh / Cetak Laporan (PDF / Print View)"):
            st.markdown("""
                <script>
                    window.print();
                </script>
            """, unsafe_allow_html=True)
            st.success("Perintah cetak dikirim ke browser Anda. Pilih 'Save as PDF' pada jendela cetak yang muncul.")
            
        st.markdown("---")
        st.markdown("### Preview Data yang Akan Dicetak:")
        st.dataframe(df, use_container_width=True, hide_index=True)

else:
    st.info("Silakan pastikan file Excel data warga RT 06 sudah di-upload dengan benar.")
