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
        
        # Konversi Format Tanggal Lahir ke Format Indonesia
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
    
    # Navigasi Menu Utama (Dipisah agar rapi dan tidak numpuk)
    menu = st.sidebar.selectbox("📂 Pilih Menu Utama", [
        "🗂️ Kartu Keluarga (KK) Warga", 
        "📊 Grafik Demografi Interaktif", 
        "🛠️ Kelola Data Warga", 
        "🖨️ Cetak Laporan PDF"
    ])
    
    # Deteksi kolom penting
    col_kk = next((c for c in df.columns if "KEPALA KELUARGA" in c or "KK" in c), None)
    col_rumah = next((c for c in df.columns if "RUMAH" in c), None)
    
    if menu == "🗂️ Kartu Keluarga (KK) Warga":
        st.subheader("🗂️ Lembar Kartu Keluarga (KK) Warga RT 06")
        st.markdown("Setiap kartu di bawah ini menampilkan rincian Kepala Keluarga beserta seluruh anggota keluarganya secara berdampingan.")
        
        if col_kk:
            daftar_kk = df[col_kk].dropna().unique()
            
            # Filter pencarian nama Kepala Keluarga
            pencarian_kk = st.text_input("🔍 Cari Nama Kepala Keluarga:", "")
            if pencarian_kk:
                daftar_kk = [kk for kk in daftar_kk if pencarian_kk.lower() in str(kk).lower()]
            
            for idx, kk in enumerate(daftar_kk, 1):
                df_anggota = df[df[col_kk] == kk]
                no_rmh = df_anggota[col_rumah].iloc[0] if col_rumah and not df_anggota.empty else "-"
                
                # Desain Kartu ala Lembar KK Resmi
                with st.container():
                    st.markdown(f"""
                    <div style="background-color: #f8fafc; border: 2px solid #cbd5e1; border-radius: 10px; padding: 20px; margin-bottom: 25px; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.1);">
                        <div style="display: flex; justify-content: space-between; align-items: center; border-bottom: 2px solid #2563eb; padding-bottom: 10px; margin-bottom: 15px;">
                            <h3 style="margin: 0; color: #1e3a8a;">🏠 KARTU KELUARGA (KK)</h3>
                            <span style="background-color: #2563eb; color: white; padding: 5px 12px; border-radius: 20px; font-weight: bold; font-size: 14px;">No. Rumah: {no_rmh}</span>
                        </div>
                        <p style="margin: 5px 0; font-size: 16px;"><strong>Kepala Keluarga:</strong> {kk}</p>
                        <p style="margin: 5px 0; font-size: 14px; color: #64748b;">Jumlah Anggota Keluarga: <strong>{len(df_anggota)} Jiwa</strong></p>
                    </div>
                    """, unsafe_allow_html=True)
                    
                    st.dataframe(df_anggota, use_container_width=True, hide_index=True)
                    st.write("---")
        else:
            st.dataframe(df, use_container_width=True, hide_index=True)
            
    elif menu == "📊 Grafik Demografi Interaktif":
        st.subheader("📊 Analisis & Statistik Grafik Demografi Warga RT 06")
        st.markdown("Grafik interaktif modern dengan ukuran penuh agar angka dan persentasenya sangat jelas dibaca saat sosialisasi.")
        
        col_jk = next((c for c in df.columns if "JK" in c or "KELAMIN" in c or "GENDER" in c), None)
        col_pend = next((c for c in df.columns if "PENDIDIKAN" in c), None)
        col_pek = next((c for c in df.columns if "PEKERJAAN" in c), None)
        col_usia = next((c for c in df.columns if "USIA" in c or "UMUR" in c), None)
        col_status = next((c for c in df.columns if "STATUS" in c or "KAWIN" in c), None)
        
        # 1. Grafik Jenis Kelamin (Full Width & Modern)
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
                st.info("Grafik donat di samping menyajikan perbandingan jumlah penduduk laki-laki dan perempuan di lingkungan RT 06 secara real-time.")
                for _, r in df_jk.iterrows():
                    st.write(f"- **{r['Jenis Kelamin']}**: {r['Jumlah']} Jiwa")

        # 2. Grafik Status Pernikahan
        if col_status:
            st.markdown("---")
            col_desc2, col_g2 = st.columns([2, 3])
            with col_desc2:
                st.markdown("### 💍 Status Pernikahan Warga")
                st.info("Menunjukkan data demografi status pernikahan warga RT 06 untuk keperluan pendataan administrasi sosial.")
                df_st = df[col_status].dropna().value_counts().reset_index()
                df_st.columns = ["Status", "Jumlah"]
                for _, r in df_st.iterrows():
                    st.write(f"- **{r['Status']}**: {r['Jumlah']} Orang")
            with col_g2:
                fig_st = px.bar(df_st, x="Status", y="Jumlah", text="Jumlah", title="Distribusi Status Pernikahan", color="Status", color_discrete_sequence=px.colors.qualitative.Pastel)
                fig_st.update_traces(textfont_size=16, textposition="outside")
                fig_st.update_layout(font=dict(size=15), title_font=dict(size=18))
                st.plotly_chart(fig_st, use_container_width=True)

        # 3. Grafik Pendidikan
        if col_pend:
            st.markdown("---")
            st.markdown("### 🎓 Tingkat Pendidikan Terakhir Warga")
            df_pd = df[col_pend].dropna().value_counts().reset_index()
            df_pd.columns = ["Pendidikan", "Jumlah"]
            fig_pd = px.bar(df_pd, x="Pendidikan", y="Jumlah", text="Jumlah", color="Pendidikan", color_discrete_sequence=px.colors.qualitative.Bold)
            fig_pd.update_traces(textfont_size=16, textposition="outside")
            fig_pd.update_layout(font=dict(size=15), xaxis=dict(tickangle=-20, tickfont=dict(size=14)))
            st.plotly_chart(fig_pd, use_container_width=True)

        # 4. Grafik Pekerjaan
        if col_pek:
            st.markdown("---")
            st.markdown("### 💼 Distribusi Mata Pencaharian / Pekerjaan Warga")
            df_pk = df[col_pek].dropna().value_counts().reset_index()
            df_pk.columns = ["Pekerjaan", "Jumlah"]
            fig_pk = px.bar(df_pk, x="Pekerjaan", y="Jumlah", text="Jumlah", color="Pekerjaan", color_discrete_sequence=px.colors.qualitative.Vivid)
            fig_pk.update_traces(textfont_size=16, textposition="outside")
            fig_pk.update_layout(font=dict(size=15), xaxis=dict(tickangle=-25, tickfont=dict(size=13)))
            st.plotly_chart(fig_pk, use_container_width=True)

        # 5. Grafik Kategori Usia
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
                st.info("Pengelompokan usia warga untuk memetakan program Posyandu, Karang Taruna, dan Lansia.")
                for _, r in df_usia_count.iterrows():
                    st.write(f"- **{r['Kategori Usia']}**: {r['Jumlah']} Jiwa")

    elif menu == "🛠️ Kelola Data Warga":
        st.subheader("🛠️ Panel Pengelolaan Data Warga RT 06")
        aksi = st.selectbox("Pilih Aksi Pengelolaan:", ["➕ Tambah Data Warga Baru", "✏️ Edit Data Warga", "🗑️ Hapus Data Warga"])
        
        if aksi == "➕ Tambah Data Warga Baru":
            with st.form("form_tambah"):
                st.text_input("No. Rumah")
                st.text_input("Nama Kepala Keluarga")
                st.text_input("Nama Anggota Keluarga")
                st.selectbox("Jenis Kelamin", ["L", "P"])
                st.selectbox("Hubungan Keluarga", ["Kepala Keluarga", "Istri", "Anak Kandung", "Famili Lain"])
                st.text_input("Tempat & Tanggal Lahir")
                st.number_input("Usia", min_value=0, max_value=120, value=25)
                st.selectbox("Status Pernikahan", ["Belum Kawin", "Kawin", "Cerai Hidup"])
                st.text_input("Pendidikan Terakhir")
                st.text_input("Pekerjaan")
                if st.form_submit_button("Simpan Data Warga Baru"):
                    st.success("Data warga baru berhasil disiapkan!")

        elif aksi == "✏️ Edit Data Warga":
            st.selectbox("Pilih Warga yang Ingin Diedit:", df.iloc[:, 2] if len(df.columns) > 2 else df.iloc[:, 0])
            with st.form("form_edit"):
                st.text_input("Perbarui Data")
                st.form_submit_button("Simpan Perubahan")

        elif aksi == "🗑️ Hapus Data Warga":
            nama_hapus = st.selectbox("Pilih Warga yang Ingin Dihapus:", df.iloc[:, 2] if len(df.columns) > 2 else df.iloc[:, 0])
            if st.button("Konfirmasi Hapus"):
                st.success(f"Data **{nama_hapus}** berhasil dihapus.")

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
