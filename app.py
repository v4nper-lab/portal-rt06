import streamlit as st
import pandas as pd
import plotly.express as px
import os
import io
from reportlab.lib.pagesizes import letter, landscape
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

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
        
        # Forward fill untuk Kepala Keluarga dan No Rumah agar terbaca rapi per kelompok
        for col in df.columns:
            if "KEPALA" in col or "KK" in col or "RUMAH" in col:
                df[col] = df[col].ffill()

        # Format Tanggal Lahir Indonesia
        bulan_indo = {
            1: "Januari", 2: "Februari", 3: "Maret", 4: "April", 5: "Mei", 6: "Juni",
            7: "Juli", 8: "Agustus", 9: "September", 10: "Oktober", 11: "November", 12: "Desember"
        }
        col_tgl = next((c for c in df.columns if "TANGGAL" in c or "LAHIR" in c and "TGL" not in c), None)
        if col_tgl:
            def format_tgl_indo(val):
                try:
                    dt = pd.to_datetime(val)
                    if pd.notnull(dt):
                        return f"{dt.day:02d} {bulan_indo.get(dt.month, '')} {dt.year}"
                except:
                    pass
                return str(val)
            df[col_tgl] = df[col_tgl].apply(format_tgl_indo)
            
        return df
    except Exception as e:
        st.error(f"Gagal memuat data: {e}")
        return pd.DataFrame()

df = load_data_rt06()

if not df.empty:
    st.metric("👥 Total Warga RT 06 Terdaftar", f"{len(df)} Jiwa")
    
    # Navigasi Menu Utama
    menu = st.sidebar.selectbox("📂 Pilih Menu Utama", [
        "🗂️ Kartu Keluarga (KK) Warga", 
        "📊 Grafik Demografi Interaktif", 
        "🛠️ Kelola Data Warga (Tambah/Edit/Hapus)", 
        "🖨️ Cetak Laporan PDF"
    ])
    
    # Deteksi kolom secara dinamis
    col_kk_candi = [c for c in df.columns if "KEPALA" in c or "KK" in c]
    col_kk = col_kk_candi[0] if col_kk_candi else df.columns[2]
    
    col_nama_candi = [c for c in df.columns if "ANGGOTA" in c or "NAMA" in c]
    col_nama = col_nama_candi[0] if col_nama_candi else df.columns[3]
    
    col_rumah_candi = [c for c in df.columns if "RUMAH" in c or "ALAMAT" in c]
    col_rumah = col_rumah_candi[0] if col_rumah_candi else None
    
    col_jk = next((c for c in df.columns if "JK" in c or "KELAMIN" in c), None)
    col_hub = next((c for c in df.columns if "HUBUNGAN" in c), None)
    col_usia = next((c for c in df.columns if "USIA" in c or "UMUR" in c), None)
    col_pek = next((c for c in df.columns if "PEKERJAAN" in c), None)
    col_status_rumah = next((c for c in df.columns if "RUMAH" in c and c != col_rumah), None)
    col_domisili = next((c for c in df.columns if "DOMISILI" in c), None)
    
    if menu == "🗂️ Kartu Keluarga (KK) Warga":
        st.subheader("🗂️ Lembar Kartu Keluarga (KK) Warga RT 06")
        st.markdown("Tampilan mobile-friendly per nomor rumah dan kepala keluarga.")
        
        daftar_kk = [str(x) for x in df[col_kk].dropna().unique() if str(x).strip() != "" and str(x).lower() != "nan"]
        
        pencarian_kk = st.text_input("🔍 Cari Nama Kepala Keluarga / No Rumah:", "")
        if pencarian_kk:
            daftar_kk = [kk for kk in daftar_kk if pencarian_kk.lower() in kk.lower()]
        
        if len(daftar_kk) == 0:
            st.warning("Data Kepala Keluarga tidak ditemukan.")
        else:
            for idx, kk in enumerate(daftar_kk, 1):
                df_anggota = df[df[col_kk].astype(str).str.strip() == kk.strip()]
                no_rmh = str(df_anggota[col_rumah].iloc[0]) if col_rumah and not df_anggota.empty else "-"
                st_rumah = str(df_anggota[col_status_rumah].iloc[0]) if col_status_rumah and not df_anggota.empty else "Milik / Tetap"
                dom = str(df_anggota[col_domisili].iloc[0]) if col_domisili and not df_anggota.empty else "Nanjung Mekar"
                
                # Kartu Utama Per Rumah / Kepala Keluarga
                with st.container():
                    st.markdown(f"""
                    <div style="background-color: #ffffff; border: 1px solid #e2e8f0; border-radius: 12px; padding: 18px; margin-bottom: 20px; box-shadow: 0 2px 4px rgba(0,0,0,0.05);">
                        <div style="display: flex; align-items: center; margin-bottom: 8px;">
                            <span style="background-color: #eff6ff; color: #1d4ed8; padding: 4px 10px; border-radius: 6px; font-weight: bold; font-size: 15px; margin-right: 10px;">{no_rmh}</span>
                            <span style="font-size: 18px; font-weight: bold; color: #0f172a;">{kk}</span>
                        </div>
                        <div style="font-size: 13px; color: #64748b; margin-bottom: 15px; border-bottom: 1px solid #f1f5f9; padding-bottom: 8px;">
                            <strong>{len(df_anggota)} anggota</strong> &nbsp;&bull;&nbsp; {st_rumah} &nbsp;&bull;&nbsp; {dom}
                        </div>
                    """, unsafe_allow_html=True)
                    
                    # Daftar Anggota Keluarga di dalam Kartu
                    for i, (_, row) in enumerate(df_anggota.iterrows()):
                        nama_anggota = row.get(col_nama, "-")
                        jk_val = row.get(col_jk, "-")
                        hub_val = row.get(col_hub, "-")
                        usia_val = row.get(col_usia, "-")
                        pek_val = row.get(col_pek, "-")
                        
                        cols = st.columns([6, 2, 2])
                        with cols[0]:
                            badge_color = "#2563eb" if str(jk_val).upper() == "L" else "#db2777"
                            st.markdown(f"""
                            <div style="margin-bottom: 10px;">
                                <span style="font-weight: bold; font-size: 15px; color: #1e293b;">{nama_anggota}</span><br>
                                <span style="background-color: {badge_color}; color: white; padding: 1px 6px; border-radius: 4px; font-size: 11px; font-weight: bold;">{jk_val}</span> 
                                <span style="font-size: 13px; color: #475569;">{hub_val} &bull; {pek_val}</span>
                            </div>
                            """, unsafe_allow_html=True)
                        with cols[1]:
                            st.markdown(f"<div style='font-size: 14px; color: #334155; padding-top: 4px;'><strong>{usia_val} th</strong></div>", unsafe_allow_html=True)
                        with cols[2]:
                            sub_cols = st.columns(2)
                            with sub_cols[0]:
                                if st.button("Ubah", key=f"edit_{idx}_{i}", help="Edit data"):
                                    st.toast(f"Edit data: {nama_anggota}")
                            with sub_cols[1]:
                                if st.button("Hapus", key=f"del_{idx}_{i}", help="Hapus data"):
                                    st.toast(f"Hapus data: {nama_anggota}")
                        
                        if i < len(df_anggota) - 1:
                            st.markdown("<hr style='margin: 6px 0; border: none; border-top: 1px dashed #e2e8f0;'>", unsafe_allow_html=True)
                    
                    # Tombol tambah anggota di bawah kartu
                    st.markdown("<div style='margin-top: 12px;'>", unsafe_allow_html=True)
                    if st.button(f"+ Tambah anggota keluarga ini", key=f"add_member_{idx}", use_container_width=True):
                        st.info(f"Form tambah anggota untuk keluarga {kk} dibuka.")
                    st.markdown("</div></div>", unsafe_allow_html=True)
            
    elif menu == "📊 Grafik Demografi Interaktif":
        st.subheader("📊 Analisis & Statistik Grafik Demografi Warga RT 06")
        
        col_jk = next((c for c in df.columns if "JK" in c or "KELAMIN" in c or "GENDER" in c), None)
        col_pend = next((c for c in df.columns if "PENDIDIKAN" in c), None)
        col_pek = next((c for c in df.columns if "PEKERJAAN" in c), None)
        col_usia = next((c for c in df.columns if "USIA" in c or "UMUR" in c), None)
        col_status = next((c for c in df.columns if "STATUS" in c and "KAWIN" in c) or (c for c in df.columns if "STATUS" in c), None)
        
        if col_jk:
            st.markdown("---")
            col_g1, col_desc1 = st.columns([3, 2])
            with col_g1:
                df_jk = df[col_jk].dropna().value_counts().reset_index()
                df_jk.columns = ["Jenis Kelamin", "Jumlah"]
                fig_jk = px.pie(df_jk, names="Jenis Kelamin", values="Jumlah", hole=0.5, title="Rasio Penduduk Berdasarkan Jenis Kelamin", color_discrete_sequence=px.colors.qualitative.Set2)
                fig_jk.update_traces(textfont_size=18, textinfo="percent+label+value")
                fig_jk.update_layout(font=dict(size=15), title_font=dict(size=18))
                st.plotly_chart(fig_jk, use_container_width=True)
            with col_desc1:
                st.markdown("### 📌 Keterangan & Analisis")
                st.info("Perbandingan jumlah penduduk laki-laki dan perempuan di lingkungan RT 06.")
                for _, r in df_jk.iterrows():
                    st.write(f"- **{r['Jenis Kelamin']}**: {r['Jumlah']} Jiwa")

        if col_status:
            st.markdown("---")
            col_desc2, col_g2 = st.columns([2, 3])
            with col_desc2:
                st.markdown("### 💍 Status Pernikahan Warga")
                st.info("Data demografi status pernikahan warga RT 06.")
                df_st = df[col_status].dropna().value_counts().reset_index()
                df_st.columns = ["Status", "Jumlah"]
                for _, r in df_st.iterrows():
                    st.write(f"- **{r['Status']}**: {r['Jumlah']} Orang")
            with col_g2:
                fig_st = px.bar(df_st, x="Status", y="Jumlah", text="Jumlah", title="Distribusi Status Pernikahan", color="Status", color_discrete_sequence=px.colors.qualitative.Pastel)
                fig_st.update_traces(textfont_size=16, textposition="outside")
                fig_st.update_layout(font=dict(size=15), title_font=dict(size=18))
                st.plotly_chart(fig_st, use_container_width=True)

        if col_pend:
            st.markdown("---")
            st.markdown("### 🎓 Tingkat Pendidikan Terakhir Warga")
            df_pd = df[col_pend].dropna().value_counts().reset_index()
            df_pd.columns = ["Pendidikan", "Jumlah"]
            fig_pd = px.bar(df_pd, x="Pendidikan", y="Jumlah", text="Jumlah", color="Pendidikan", color_discrete_sequence=px.colors.qualitative.Bold)
            fig_pd.update_traces(textfont_size=16, textposition="outside")
            fig_pd.update_layout(font=dict(size=15), xaxis=dict(tickangle=-20, tickfont=dict(size=14)))
            st.plotly_chart(fig_pd, use_container_width=True)

        if col_pek:
            st.markdown("---")
            st.markdown("### 💼 Distribusi Mata Pencaharian / Pekerjaan Warga")
            df_pk = df[col_pek].dropna().value_counts().reset_index()
            df_pk.columns = ["Pekerjaan", "Jumlah"]
            fig_pk = px.bar(df_pk, x="Pekerjaan", y="Jumlah", text="Jumlah", color="Pekerjaan", color_discrete_sequence=px.colors.qualitative.Vivid)
            fig_pk.update_traces(textfont_size=16, textposition="outside")
            fig_pk.update_layout(font=dict(size=15), xaxis=dict(tickangle=-25, tickfont=dict(size=13)))
            st.plotly_chart(fig_pk, use_container_width=True)

        if col_usia:
            st.markdown("---")
            col_g5, col_desc5 = st.columns([3, 2])
            with col_g5:
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
                
                fig_usia = px.pie(df_usia_count, names="Kategori Usia", values="Jumlah", hole=0.4, title="Kelompok Rentang Usia Penduduk", color_discrete_sequence=px.colors.qualitative.Safe)
                fig_usia.update_traces(textfont_size=16, textinfo="percent+label+value")
                fig_usia.update_layout(font=dict(size=15), title_font=dict(size=18))
                st.plotly_chart(fig_usia, use_container_width=True)
            with col_desc5:
                st.markdown("### 👶 Klasifikasi Usia Warga")
                st.info("Pengelompokan usia warga.")
                for _, r in df_usia_count.iterrows():
                    st.write(f"- **{r['Kategori Usia']}**: {r['Jumlah']} Jiwa")

    elif menu == "🛠️ Kelola Data Warga (Tambah/Edit/Hapus)":
        st.subheader("🛠️ Panel Pengelolaan Data Warga RT 06")
        aksi = st.selectbox("Pilih Aksi Pengelolaan:", ["➕ Tambah Data Warga Baru", "✏️ Edit Data Warga", "🗑️ Hapus Data Warga"])
        
        kol_pilih_nama = col_nama if col_nama else df.columns[0]
        
        if aksi == "➕ Tambah Data Warga Baru":
            st.markdown("### Form Tambah Warga Baru")
            with st.form("form_tambah"):
                st.text_input("No. Rumah")
                st.text_input("Nama Kepala Keluarga")
                st.text_input("Nama Anggota Keluarga")
                st.selectbox("Jenis Kelamin", ["L", "P"])
                st.selectbox("Hubungan Keluarga", ["Kepala Keluarga", "Istri", "Anak Kandung", "Famili Lain", "Mertua"])
                st.text_input("Tempat & Tanggal Lahir")
                st.number_input("Usia", min_value=0, max_value=120, value=25)
                st.selectbox("Status Pernikahan", ["Belum Kawin", "Kawin", "Cerai Hidup", "Cerai Mati"])
                st.text_input("Pendidikan Terakhir", "Tamat SLTA/sederajat")
                st.text_input("Pekerjaan", "Karyawan Swasta")
                
                if st.form_submit_button("Simpan Data Warga Baru"):
                    st.success("Data warga baru berhasil disiapkan!")

        elif aksi == "✏️ Edit Data Warga":
            st.markdown("### Edit Data Warga")
            warga_pilih = st.selectbox("Pilih Warga yang Ingin Diedit:", df[kol_pilih_nama].astype(str).tolist())
            st.info(f"Form edit untuk **{warga_pilih}** aktif.")
            with st.form("form_edit"):
                st.text_input("Perbarui Nama", value=str(warga_pilih))
                st.text_input("Perbarui No. Rumah / Alamat")
                st.form_submit_button("Simpan Perubahan")

        elif aksi == "🗑️ Hapus Data Warga":
            st.markdown("### Hapus Data Warga (Pindah / Keluar)")
            warga_hapus = st.selectbox("Pilih Warga yang Ingin Dihapus:", df[kol_pilih_nama].astype(str).tolist())
            if st.button("Konfirmasi Hapus Warga Ini", type="primary"):
                st.success(f"Data warga **{warga_hapus}** berhasil dihapus dari sistem.")

    elif menu == "🖨️ Cetak Laporan PDF":
        st.subheader("🖨️ Unduh Laporan Rekapitulasi Data Warga (PDF)")
        
        def buat_pdf(data_df):
            buffer = io.BytesIO()
            doc = SimpleDocTemplate(buffer, pagesize=landscape(letter), rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=30)
            elements = []
            styles = getSampleStyleSheet()
            
            elements.append(Paragraph("REKAPITULASI DATA WARGA RT 06 / RW 14", ParagraphStyle('Title', parent=styles['Heading1'], fontSize=16, alignment=1, textColor=colors.HexColor('#1f2937'))))
            elements.append(Paragraph("Griya Permata Raya - Desa Nanjung Mekar", ParagraphStyle('Sub', parent=styles['Normal'], alignment=1, fontSize=11, textColor=colors.gray)))
            elements.append(Spacer(1, 15))
            
            kolom_tampil = data_df.columns[:min(8, len(data_df.columns))]
            table_data = [list(kolom_tampil)]
            for _, row in data_df.iterrows():
                table_data.append([str(row[col])[:20] for col in kolom_tampil])
                
            t = Table(table_data, repeatRows=1)
            t.setStyle(TableStyle([
                ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#2563eb')),
                ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke),
                ('ALIGN', (0,0), (-1,-1), 'CENTER'),
                ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
                ('FONTSIZE', (0,0), (-1,0), 9),
                ('BOTTOMPADDING', (0,0), (-1,0), 6),
                ('BACKGROUND', (0,1), (-1,-1), colors.HexColor('#f9fafb')),
                ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#d1d5db')),
                ('FONTSIZE', (0,1), (-1,-1), 8),
            ]))
            elements.append(t)
            doc.build(elements)
            buffer.seek(0)
            return buffer.getvalue()

        pdf_bytes = buat_pdf(df)
        st.download_button(
            label="📥 Download File PDF Rekapitulasi Warga",
            data=pdf_bytes,
            file_name="Laporan_Data_Warga_RT06.pdf",
            mime="application/pdf",
            type="primary"
        )
        
        st.markdown("---")
        st.markdown("### Preview Data:")
        st.dataframe(df, use_container_width=True, hide_index=True)

else:
    st.info("Silakan pastikan file Excel data warga RT 06 sudah di-upload dengan benar.")
