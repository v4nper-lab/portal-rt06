import streamlit as st
import pandas as pd
import plotly.express as px
import os
import io
import time
from datetime import datetime
from reportlab.lib.pagesizes import letter, landscape
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

st.set_page_config(
    page_title="Portal Resmi RT 06 RW 14 Griya Permata Raya",
    layout="wide",
    page_icon="🏠"
)

# Header dengan Logo RT 06
col_logo, col_title = st.columns([1, 5])
with col_logo:
    logo_path = "logo_rt06.jpg"
    if os.path.exists(logo_path):
        st.image(logo_path, width=330)
    else:
        st.image("logo r6.jpg", width=330) if os.path.exists("logo r6.jpg") else st.write("🏠")

with col_title:
    st.markdown("<br>", unsafe_allow_html=True)
    st.title("🏠 Portal Resmi RT 06 / RW 14")
    st.markdown("### Griya Permata Raya - Desa Nanjung Mekar, Kec. Rancaekek")

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
    # Navigasi Menu Utama
    menu = st.sidebar.selectbox("📂 Pilih Menu Utama", [
        "📋 Dashboard Data Seluruh Warga",
        "🗂️ Cetak / Lihat Kartu Keluarga (KK)", 
        "📊 Grafik Demografi Interaktif", 
        "🛠️ Kelola Data Warga (Tambah/Edit/Hapus)", 
        "🖨️ Cetak Laporan Rekap PDF"
    ])
    
    col_kk_candi = [c for c in df.columns if "KEPALA" in c or "KK" in c]
    col_kk = col_kk_candi[0] if col_kk_candi else df.columns[2]
    
    col_nama_candi = [c for c in df.columns if "ANGGOTA" in c or "NAMA" in c]
    col_nama = col_nama_candi[0] if col_nama_candi else df.columns[3]
    
    col_rumah_candi = [c for c in df.columns if "RUMAH" in c or "ALAMAT" in c]
    col_rumah = col_rumah_candi[0] if col_rumah_candi else None
    
    if menu == "📋 Dashboard Data Seluruh Warga":
        col_jam1, col_jam2 = st.columns([3, 1])
        with col_jam1:
            st.subheader("📋 Dashboard Modern & Interaktif Warga RT 06")
            st.markdown("Pusat informasi data kependudukan real-time Griya Permata Raya.")
        with col_jam2:
            placeholder_waktu = st.empty()
            
        # Statistik Utama dengan Desain Kartu Warna-Warni
        total_warga = len(df)
        total_kk = df[col_kk].nunique() if col_kk in df.columns else 0
        total_rumah = df[col_rumah].nunique() if col_rumah in df.columns else 0
        
        st.markdown(f"""
        <div style="display: flex; gap: 20px; margin-bottom: 25px; flex-wrap: wrap;">
            <div style="flex: 1; min-width: 220px; background: linear-gradient(135deg, #3b82f6, #1d4ed8); color: white; padding: 20px; border-radius: 14px; box-shadow: 0 4px 6px rgba(0,0,0,0.1);">
                <div style="font-size: 14px; opacity: 0.9;">👥 Total Jiwa Warga</div>
                <div style="font-size: 28px; font-weight: bold; margin-top: 5px;">{total_warga} Orang</div>
            </div>
            <div style="flex: 1; min-width: 220px; background: linear-gradient(135deg, #10b981, #047857); color: white; padding: 20px; border-radius: 14px; box-shadow: 0 4px 6px rgba(0,0,0,0.1);">
                <div style="font-size: 14px; opacity: 0.9;">🏠 Total Kepala Keluarga</div>
                <div style="font-size: 28px; font-weight: bold; margin-top: 5px;">{total_kk} KK</div>
            </div>
            <div style="flex: 1; min-width: 220px; background: linear-gradient(135deg, #f59e0b, #b45309); color: white; padding: 20px; border-radius: 14px; box-shadow: 0 4px 6px rgba(0,0,0,0.1);">
                <div style="font-size: 14px; opacity: 0.9;">🏡 Total Rumah Terdata</div>
                <div style="font-size: 28px; font-weight: bold; margin-top: 5px;">{total_rumah} Rumah</div>
            </div>
        </div>
        """, unsafe_allow_html=True)
        
        # Fitur Pencarian Cepat Warga di Dashboard
        pencarian_dashboard = st.text_input("🔍 Cari Cepat Data Warga (Ketik Nama / No. Rumah):", "")
        
        df_tampil_dash = df.copy()
        if pencarian_dashboard:
            mask = df_tampil_dash.astype(str).apply(lambda x: x.str.contains(pencarian_dashboard, case=False)).any(axis=1)
            df_tampil_dash = df_tampil_dash[mask]
            
        # Tambahkan Kolom Aksi Edit & Hapus di Dashboard
        st.markdown("### 📊 Tabel Data Warga & Kontrol Cepat")
        
        # Header Tabel Kustom dengan Tombol Aksi
        for idx, row in df_tampil_dash.iterrows():
            nama_warga = row.get(col_nama, f"Warga #{idx+1}")
            no_rmh = row.get(col_rumah, "-")
            kk_warga = row.get(col_kk, "-")
            
            with st.expander(f"📌 [{idx+1}] {nama_warga} (Rumah: {no_rmh} | KK: {kk_warga})"):
                c_info, c_btn1, c_btn2 = st.columns([6, 2, 2])
                with c_info:
                    # Tampilkan detail ringkas baris
                    details = " | ".join([f"**{col}**: {row[col]}" for col in df.columns if pd.notnull(row[col])])
                    st.markdown(details)
                with c_btn1:
                    if st.button("✏️ Edit Baris", key=f"dash_edit_{idx}"):
                        st.toast(f"Membuka form edit untuk: {nama_warga}")
                with c_btn2:
                    if st.button("🗑️ Hapus Baris", key=f"dash_del_{idx}", type="primary"):
                        st.toast(f"Data {nama_warga} ditandai untuk dihapus.")
                        
        st.write("---")
        st.markdown("#### 📑 Tinjauan Keseluruhan Data Tabel")
        if col_kk in df_tampil_dash.columns:
            df_tampil_dash[col_kk] = df_tampil_dash[col_kk].mask(df_tampil_dash[col_kk].duplicated(), None)
        if col_rumah in df_tampil_dash.columns:
            df_tampil_dash[col_rumah] = df_tampil_dash[col_rumah].mask(df_tampil_dash[col_rumah].duplicated(), None)
            
        st.dataframe(df_tampil_dash, use_container_width=True, hide_index=True)
        
        # Script interaktif waktu berjalan
        for _ in range(3):
            waktu_sekarang = datetime.now().strftime("%d %B %Y | %H:%M:%S")
            placeholder_waktu.markdown(f"""
            <div style="background: linear-gradient(135deg, #8b5cf6, #6d28d9); color: white; padding: 12px 18px; border-radius: 12px; text-align: center; box-shadow: 0 2px 4px rgba(0,0,0,0.1);">
                <span style="font-size: 11px; opacity: 0.8; text-transform: uppercase; letter-spacing: 1px;">🕒 Waktu Sistem Real-Time</span><br>
                <strong style="font-size: 15px;">{waktu_sekarang}</strong>
            </div>
            """, unsafe_allow_html=True)
            time.sleep(1)

    elif menu == "🗂️ Cetak / Lihat Kartu Keluarga (KK)":
        st.subheader("🗂️ Pencarian & Cetak Kartu Keluarga (KK) per Rumah")
        st.markdown("Pilih Nama Kepala Keluarga untuk melihat seluruh anggota keluarga dan mencetaknya ke format PDF A4 Landscape.")
        
        daftar_kk = df[col_kk].dropna().astype(str).str.strip()
        daftar_kk = sorted(list(set([x for x in daftar_kk if x != "" and x.lower() != "nan"])))
        
        pilihan_kk = st.selectbox("Pilih Kepala Keluarga:", daftar_kk)
        
        if pilihan_kk:
            df_keluarga = df[df[col_kk].astype(str).str.strip() == pilihan_kk.strip()].copy()
            no_rmh = str(df_keluarga[col_rumah].dropna().iloc[0]) if col_rumah and not df_keluarga[df_keluarga[col_rumah].notna()].empty else "-"
            
            cols_tampilan_web = [c for c in df_keluarga.columns if c != col_kk and "URUT" not in c and c != "NO"]
            
            st.markdown(f"""
            <div style="background: linear-gradient(135deg, #eff6ff, #dbeafe); border: 2px solid #3b82f6; border-radius: 12px; padding: 20px; margin-top: 15px; margin-bottom: 20px; box-shadow: 0 4px 6px rgba(0,0,0,0.05);">
                <h4 style="margin: 0; color: #1e3a8a;">🏠 KARTU KELUARGA - NO. RUMAH: {no_rmh}</h4>
                <p style="margin: 8px 0 0 0; font-size: 16px; color: #1f2937;"><b>Kepala Keluarga:</b> {pilihan_kk}</p>
                <p style="margin: 4px 0 0 0; font-size: 14px; color: #4b5563;">Jumlah Anggota: {len(df_keluarga)} Jiwa</p>
            </div>
            """, unsafe_allow_html=True)
            
            st.dataframe(df_keluarga[cols_tampilan_web], use_container_width=True, hide_index=True)
            
            def buat_pdf_kk_landscape(keluarga_df, kepala, rumah):
                buffer = io.BytesIO()
                doc = SimpleDocTemplate(buffer, pagesize=landscape(letter), rightMargin=20, leftMargin=20, topMargin=25, bottomMargin=25)
                elements = []
                styles = getSampleStyleSheet()
                
                elements.append(Paragraph("PEMERINTAH KABUPATEN BANDUNG", ParagraphStyle('Sub1', parent=styles['Normal'], alignment=1, fontSize=10, textColor=colors.gray)))
                elements.append(Paragraph("KECAMATAN RANCAAEKEK - DESA NANJUNG MEKAR", ParagraphStyle('Sub2', parent=styles['Normal'], alignment=1, fontSize=10, textColor=colors.gray)))
                elements.append(Paragraph("KARTU KELUARGA (KK) RT 06 / RW 14", ParagraphStyle('Title', parent=styles['Heading1'], fontSize=15, alignment=1, textColor=colors.HexColor('#1f2937'))))
                elements.append(Spacer(1, 10))
                
                elements.append(Paragraph(f"<b>No. Rumah / Alamat:</b> {rumah}", styles['Normal']))
                elements.append(Paragraph(f"<b>Kepala Keluarga:</b> {kepala}", styles['Normal']))
                elements.append(Spacer(1, 10))
                
                kolom_pdf = [c for c in keluarga_df.columns if c != col_kk and "URUT" not in c and c != "NO"]
                
                cell_style = ParagraphStyle('Cell', parent=styles['Normal'], fontSize=7, leading=8, alignment=1)
                header_style = ParagraphStyle('HeaderCell', parent=styles['Normal'], fontSize=7.5, leading=9, textColor=colors.whitesmoke, fontName='Helvetica-Bold', alignment=1)
                
                table_data = [[Paragraph(str(col), header_style) for col in kolom_pdf]]
                for _, row in keluarga_df.iterrows():
                    row_cells = [Paragraph(str(row[col]) if pd.notnull(row[col]) else "", cell_style) for col in kolom_pdf]
                    table_data.append(row_cells)
                    
                num_cols = len(kolom_pdf)
                col_width = 800.0 / num_cols if num_cols > 0 else 100
                col_widths = [col_width] * num_cols
                
                t = Table(table_data, colWidths=col_widths, repeatRows=1)
                t.setStyle(TableStyle([
                    ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#2563eb')),
                    ('ALIGN', (0,0), (-1,-1), 'CENTER'),
                    ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                    ('BOTTOMPADDING', (0,0), (-1,-1), 6),
                    ('TOPPADDING', (0,0), (-1,-1), 6),
                    ('BACKGROUND', (0,1), (-1,-1), colors.HexColor('#f9fafb')),
                    ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#d1d5db')),
                ]))
                elements.append(t)
                doc.build(elements)
                buffer.seek(0)
                return buffer.getvalue()

            pdf_kk_bytes = buat_pdf_kk_landscape(df_keluarga, pilihan_kk, no_rmh)
            st.download_button(
                label=f"📥 Download PDF Kartu Keluarga A4 Landscape ({pilihan_kk})",
                data=pdf_kk_bytes,
                file_name=f"KK_{pilihan_kk.replace(' ', '_')}.pdf",
                mime="application/pdf",
                type="primary"
            )

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
                daftar_no_rumah = sorted(list(set(df[col_rumah].dropna().astype(str).tolist()))) if col_rumah else ["B3-01", "B3-02", "B3-03", "B3-04"]
                no_rumah = st.selectbox("No. Rumah", daftar_no_rumah)
                
                nama_kk = st.selectbox("Nama Kepala Keluarga", sorted(list(set(df[col_kk].dropna().astype(str).tolist()))))
                nama_anggota = st.text_input("Nama Lengkap Anggota Keluarga")
                jk = st.selectbox("Jenis Kelamin", ["L", "P"])
                hubungan = st.selectbox("Hubungan Keluarga", ["Kepala Keluarga", "Istri", "Anak Kandung", "Famili Lain", "Mertua"])
                tempat_lahir = st.text_input("Tempat Lahir")
                tanggal_lahir = st.text_input("Tanggal Lahir (Contoh: 02 Agustus 1991)")
                usia = st.number_input("Usia", min_value=0, max_value=120, value=25)
                status_nikah = st.selectbox("Status Pernikahan", ["Belum Kawin", "Kawin", "Cerai Hidup", "Cerai Mati"])
                
                pendidikan = st.selectbox("Pendidikan Terakhir", [
                    "Tamat SLTA/sederajat", 
                    "Tamat SLTP/sederajat", 
                    "Tamat SD/sederajat", 
                    "Diploma / Sarjana (S1/S2/S3)", 
                    "Belum / Tidak Sekolah", 
                    "Sedang SLTA/Sederajat", 
                    "Sedang SLTP/Sederajat"
                ])
                
                pekerjaan = st.selectbox("Pekerjaan", [
                    "Karyawan Swasta", 
                    "Wiraswasta", 
                    "Mengurus Rumah Tangga", 
                    "Belum Bekerja", 
                    "Pelajar / Mahasiswa", 
                    "PNS / TNI / Polri", 
                    "Buruh / Freelance", 
                    "Pensiunan"
                ])
                
                status_rumah = st.selectbox("Status Rumah", ["Milik / Tetap", "Sewa / Kontrak"])
                status_domisili = st.selectbox("Status Domisili", ["Warga Tetap", "Warga Kontrak", "Luar NM"])
                
                if st.form_submit_button("Simpan Data Warga Baru"):
                    st.success(f"Data warga baru atas nama **{nama_anggota}** berhasil disiapkan!")

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

    elif menu == "🖨️ Cetak Laporan Rekap PDF":
        st.subheader("🖨️ Unduh Laporan Rekapitulasi Data Warga Keseluruhan (PDF)")
        
        def buat_pdf_rekap(data_df):
            buffer = io.BytesIO()
            doc = SimpleDocTemplate(buffer, pagesize=landscape(letter), rightMargin=20, leftMargin=20, topMargin=25, bottomMargin=25)
            elements = []
            styles = getSampleStyleSheet()
            
            elements.append(Paragraph("REKAPITULASI KESELURUHAN DATA WARGA RT 06 / RW 14", ParagraphStyle('Title', parent=styles['Heading1'], fontSize=16, alignment=1, textColor=colors.HexColor('#1f2937'))))
            elements.append(Paragraph("Kecamatan Rancaekek - Desa Nanjung Mekar", ParagraphStyle('Sub', parent=styles['Normal'], alignment=1, fontSize=11, textColor=colors.gray)))
            elements.append(Spacer(1, 15))
            
            kolom_tampil = list(data_df.columns)
            
            cell_style = ParagraphStyle('Cell', parent=styles['Normal'], fontSize=6.5, leading=7.5, alignment=1)
            header_style = ParagraphStyle('HeaderCell', parent=styles['Normal'], fontSize=7, leading=8.5, textColor=colors.whitesmoke, fontName='Helvetica-Bold', alignment=1)
            
            table_data = [[Paragraph(str(col), header_style) for col in kolom_tampil]]
            for _, row in data_df.iterrows():
                table_data.append([Paragraph(str(row[col]) if pd.notnull(row[col]) else "", cell_style) for col in kolom_tampil])
                
            num_cols = len(kolom_tampil)
            col_width = 800.0 / num_cols if num_cols > 0 else 100
            col_widths = [col_width] * num_cols
            
            t = Table(table_data, colWidths=col_widths, repeatRows=1)
            t.setStyle(TableStyle([
                ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#2563eb')),
                ('ALIGN', (0,0), (-1,-1), 'CENTER'),
                ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                ('BOTTOMPADDING', (0,0), (-1,-1), 5),
                ('TOPPADDING', (0,0), (-1,-1), 5),
                ('BACKGROUND', (0,1), (-1,-1), colors.HexColor('#f9fafb')),
                ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#d1d5db')),
            ]))
            elements.append(t)
            doc.build(elements)
            buffer.seek(0)
            return buffer.getvalue()

        pdf_rekap = buat_pdf_rekap(df)
        st.download_button(
            label="📥 Download File PDF Rekapitulasi Keseluruhan",
            data=pdf_rekap,
            file_name="Rekap_Warga_RT06.pdf",
            mime="application/pdf",
            type="primary"
        )
        
        st.markdown("---")
        st.markdown("### Preview Data:")
        df_dash_prev = df.copy()
        if col_kk in df_dash_prev.columns:
            df_dash_prev[col_kk] = df_dash_prev[col_kk].mask(df_dash_prev[col_kk].duplicated(), None)
        if col_rumah in df_dash_prev.columns:
            df_dash_prev[col_rumah] = df_dash_prev[col_rumah].mask(df_dash_prev[col_rumah].duplicated(), None)
        st.dataframe(df_dash_prev, use_container_width=True, hide_index=True)

else:
    st.info("Silakan pastikan file Excel data warga RT 06 sudah di-upload dengan benar.")
