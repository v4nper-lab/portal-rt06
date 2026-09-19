import streamlit as st
import pandas as pd
import plotly.express as px
import os
import io
import time
from datetime import datetime, date
from reportlab.lib.pagesizes import letter, landscape
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

st.set_page_config(
    page_title="Portal Resmi RT 06 RW 14 Griya Permata Raya",
    layout="wide",
    page_icon="🏠"
)

# Custom CSS untuk background elegan dan modern
st.markdown("""
<style>
    .stApp {
        background: linear-gradient(135deg, #f8fafc 0%, #e2e8f0 100%);
    }
    .metric-card {
        background: rgba(255, 255, 255, 0.85);
        backdrop-filter: blur(10px);
        border: 1px solid rgba(226, 232, 240, 0.8);
        padding: 24px;
        border-radius: 16px;
        box-shadow: 0 10px 25px -5px rgba(0, 0, 0, 0.05), 0 8px 10px -6px rgba(0, 0, 0, 0.05);
        transition: transform 0.2s ease, box-shadow 0.2s ease;
    }
    .metric-card:hover {
        transform: translateY(-3px);
        box-shadow: 0 20px 25px -5px rgba(0, 0, 0, 0.1), 0 10px 10px -5px rgba(0, 0, 0, 0.04);
    }
</style>
""", unsafe_allow_html=True)

# Header Utama Portal RT 06 dengan Logo
col_logo, col_title = st.columns([1, 5])
with col_logo:
    logo_path = "logo_rt06.jpg"
    if os.path.exists(logo_path):
        st.image(logo_path, width=330)
    else:
        st.image("logo r6.jpg", width=330) if os.path.exists("logo r6.jpg") else st.write("🏠")

with col_title:
    st.markdown("<br>", unsafe_allow_html=True)
    st.title("🏠 PORTAL RESMI RT 06 / RW 14")
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
        
        for c in df.columns:
            if "TGL" in c or "TANGGAL" in c or "LAHIR" in c:
                def format_tgl_indo(val):
                    try:
                        dt = pd.to_datetime(val, errors='coerce')
                        if pd.notnull(dt):
                            return f"{dt.day:02d} {bulan_indo.get(dt.month, '')} {dt.year}"
                    except:
                        pass
                    val_str = str(val)
                    if "00:00:00" in val_str:
                        val_str = val_str.replace("00:00:00", "").strip()
                    return val_str
                df[c] = df[c].apply(format_tgl_indo)
            
        return df
    except Exception as e:
        st.error(f"Gagal memuat data: {e}")
        return pd.DataFrame()

if 'df_warga' not in st.session_state:
    st.session_state.df_warga = load_data_rt06()

df = st.session_state.df_warga

if not df.empty:
    menu = st.sidebar.selectbox("📂 Navigasi Menu", [
        "📊 Dashboard & Rekapitulasi",
        "📋 Data Seluruh Warga",
        "🗂️ Cetak Kartu Keluarga (KK)", 
        "📈 Grafik Demografi", 
        "🛠️ Kelola Warga", 
        "🖨️ Cetak Rekap PDF"
    ])
    
    col_kk_candi = [c for c in df.columns if "KEPALA" in c or "KK" in c]
    col_kk = col_kk_candi[0] if col_kk_candi else df.columns[2]
    
    col_nama_candi = [c for c in df.columns if "ANGGOTA" in c or "NAMA" in c]
    col_nama = col_nama_candi[0] if col_nama_candi else df.columns[3]
    
    col_rumah_candi = [c for c in df.columns if "RUMAH" in c or "ALAMAT" in c]
    col_rumah = col_rumah_candi[0] if col_rumah_candi else None

    col_jk = next((c for c in df.columns if "JK" in c or "KELAMIN" in c or "GENDER" in c), None)
    col_usia = next((c for c in df.columns if "USIA" in c or "UMUR" in c), None)

    if menu == "📊 Dashboard & Rekapitulasi":
        col_jam1, col_jam2 = st.columns([3, 1])
        with col_jam1:
            st.subheader("📊 Dashboard Eksekutif Kependudukan RT 06")
        with col_jam2:
            placeholder_waktu = st.empty()

        total_jiwa = len(df)
        total_kk = df[col_kk].nunique() if col_kk in df.columns else 0
        
        jml_l = 0
        jml_p = 0
        if col_jk:
            jml_l = len(df[df[col_jk].astype(str).str.upper().str.contains("L")])
            jml_p = len(df[df[col_jk].astype(str).str.upper().str.contains("P")])

        jml_balita = 0
        jml_lansia = 0
        if col_usia:
            def hitung_kategori(u):
                try:
                    return int(u)
                except:
                    return -1
            usia_series = df[col_usia].apply(hitung_kategori)
            jml_balita = len(usia_series[(usia_series >= 0) & (usia_series <= 5)])
            jml_lansia = len(usia_series[usia_series > 60])

        st.markdown(f"""
        <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(240px, 1fr)); gap: 20px; margin-top: 15px; margin-bottom: 30px;">
            <div class="metric-card" style="border-left: 5px solid #2563eb;">
                <div style="font-size: 12px; text-transform: uppercase; font-weight: 600; color: #64748b; letter-spacing: 0.5px;">Jumlah KK (Kepala Keluarga)</div>
                <div style="font-size: 32px; font-weight: 800; color: #1e3a8a; margin-top: 8px;">{total_kk} <span style="font-size: 16px; font-weight: 500; color: #64748b;">KK</span></div>
            </div>
            <div class="metric-card" style="border-left: 5px solid #059669;">
                <div style="font-size: 12px; text-transform: uppercase; font-weight: 600; color: #64748b; letter-spacing: 0.5px;">Jumlah Jiwa Total</div>
                <div style="font-size: 32px; font-weight: 800; color: #065f46; margin-top: 8px;">{total_jiwa} <span style="font-size: 16px; font-weight: 500; color: #64748b;">Jiwa</span></div>
            </div>
            <div class="metric-card" style="border-left: 5px solid #0284c7;">
                <div style="font-size: 12px; text-transform: uppercase; font-weight: 600; color: #64748b; letter-spacing: 0.5px;">Laki-laki</div>
                <div style="font-size: 32px; font-weight: 800; color: #0369a1; margin-top: 8px;">{jml_l} <span style="font-size: 16px; font-weight: 500; color: #64748b;">Orang</span></div>
            </div>
            <div class="metric-card" style="border-left: 5px solid #db2777;">
                <div style="font-size: 12px; text-transform: uppercase; font-weight: 600; color: #64748b; letter-spacing: 0.5px;">Perempuan</div>
                <div style="font-size: 32px; font-weight: 800; color: #9d174d; margin-top: 8px;">{jml_p} <span style="font-size: 16px; font-weight: 500; color: #64748b;">Orang</span></div>
            </div>
            <div class="metric-card" style="border-left: 5px solid #d97706;">
                <div style="font-size: 12px; text-transform: uppercase; font-weight: 600; color: #64748b; letter-spacing: 0.5px;">Jumlah Balita (0-5 tahun)</div>
                <div style="font-size: 32px; font-weight: 800; color: #b45309; margin-top: 8px;">{jml_balita} <span style="font-size: 16px; font-weight: 500; color: #64748b;">Jiwa</span></div>
            </div>
            <div class="metric-card" style="border-left: 5px solid #7c3aed;">
                <div style="font-size: 12px; text-transform: uppercase; font-weight: 600; color: #64748b; letter-spacing: 0.5px;">Jumlah Lansia (>60 tahun)</div>
                <div style="font-size: 32px; font-weight: 800; color: #5b21b6; margin-top: 8px;">{jml_lansia} <span style="font-size: 16px; font-weight: 500; color: #64748b;">Jiwa</span></div>
            </div>
        </div>
        """, unsafe_allow_html=True)

        for _ in range(3):
            waktu_sekarang = datetime.now().strftime("%d %B %Y | %H:%M:%S")
            placeholder_waktu.markdown(f"""
            <div style="background: rgba(255, 255, 255, 0.9); border: 1px solid #cbd5e1; padding: 10px 15px; border-radius: 12px; text-align: right; box-shadow: 0 4px 6px rgba(0,0,0,0.02);">
                <span style="font-size: 11px; color: #64748b;">🕒 Waktu Sistem:</span><br>
                <strong style="font-size: 13px; color: #0f172a;">{waktu_sekarang}</strong>
            </div>
            """, unsafe_allow_html=True)
            time.sleep(1)

    elif menu == "📋 Data Seluruh Warga":
        st.subheader("📋 Data Keseluruhan Warga RT 06")
        st.data_editor(df, num_rows="dynamic", use_container_width=True, key="editor_warga_grid")

    elif menu == "🗂️ Cetak Kartu Keluarga (KK)":
        st.subheader("🗂️ Cetak Kartu Keluarga (KK) per Rumah")
        daftar_kk = df[col_kk].dropna().astype(str).str.strip()
        daftar_kk = sorted(list(set([x for x in daftar_kk if x != "" and x.lower() != "nan"])))
        
        pilihan_kk = st.selectbox("Pilih Kepala Keluarga:", daftar_kk)
        
        if pilihan_kk:
            df_keluarga = df[df[col_kk].astype(str).str.strip() == pilihan_kk.strip()].copy()
            no_rmh = str(df_keluarga[col_rumah].dropna().iloc[0]) if col_rumah and not df_keluarga[df_keluarga[col_rumah].notna()].empty else "-"
            
            cols_tampilan_web = [c for c in df_keluarga.columns if c != col_kk and "URUT" not in c and c != "NO"]
            
            st.markdown(f"""
            <div style="background: #ffffff; border: 2px solid #2563eb; border-radius: 12px; padding: 15px; margin-bottom: 15px; box-shadow: 0 4px 6px rgba(0,0,0,0.05);">
                <h4 style="margin: 0; color: #1e3a8a;">🏠 No. Rumah: {no_rmh} | Kepala Keluarga: {pilihan_kk}</h4>
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
                label=f"📥 Download PDF Kartu Keluarga ({pilihan_kk})",
                data=pdf_kk_bytes,
                file_name=f"KK_{pilihan_kk.replace(' ', '_')}.pdf",
                mime="application/pdf",
                type="primary"
            )

    elif menu == "📈 Grafik Demografi":
        st.subheader("📈 Analisis Grafik Demografi Warga")
        if col_jk:
            df_jk = df[col_jk].dropna().value_counts().reset_index()
            df_jk.columns = ["Jenis Kelamin", "Jumlah"]
            fig_jk = px.pie(df_jk, names="Jenis Kelamin", values="Jumlah", hole=0.4, title="Rasio Jenis Kelamin", color_discrete_sequence=px.colors.qualitative.Set2)
            st.plotly_chart(fig_jk, use_container_width=True)

    elif menu == "🛠️ Kelola Warga":
        st.subheader("🛠️ Panel Pengelolaan Data Warga")
        aksi = st.selectbox("Pilih Aksi Pengelolaan:", ["➕ Tambah Data Warga Baru"])
        
        if aksi == "➕ Tambah Data Warga Baru":
            st.markdown("### Form Tambah Warga Baru")
            with st.form("form_tambah"):
                daftar_no_rumah = sorted(list(set(df[col_rumah].dropna().astype(str).tolist()))) if col_rumah else ["B3-01", "B3-02", "B3-03"]
                no_rumah = st.selectbox("No. Rumah", daftar_no_rumah)
                
                nama_kk = st.selectbox("Nama Kepala Keluarga", sorted(list(set(df[col_kk].dropna().astype(str).tolist()))))
                nama_anggota = st.text_input("Nama Lengkap Anggota Keluarga")
                jk = st.selectbox("Jenis Kelamin", ["L", "P"])
                hubungan = st.selectbox("Hubungan Keluarga", ["Kepala Keluarga", "Istri", "Anak Kandung", "Famili Lain", "Mertua"])
                tempat_lahir = st.text_input("Tempat Lahir")
                
                # Date Picker Kalender Interaktif
                tanggal_lahir_date = st.date_input("Tanggal Lahir", value=date(1995, 1, 1))
                bulan_indo_nama = {1: "Januari", 2: "Februari", 3: "Maret", 4: "April", 5: "Mei", 6: "Juni", 7: "Juli", 8: "Agustus", 9: "September", 10: "Oktober", 11: "November", 12: "Desember"}
                tanggal_lahir_str = f"{tanggal_lahir_date.day:02d} {bulan_indo_nama.get(tanggal_lahir_date.month, '')} {tanggal_lahir_date.year}"
                
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
                    st.success(f"Data warga baru atas nama **{nama_anggota}** (Lahir: {tanggal_lahir_str}) berhasil disiapkan!")

    elif menu == "🖨️ Cetak Rekap PDF":
        st.subheader("🖨️ Cetak Laporan Rekapitulasi Keseluruhan (PDF)")
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
            label="📥 Download PDF Rekapitulasi Keseluruhan",
            data=pdf_rekap,
            file_name="Rekap_Warga_RT06.pdf",
            mime="application/pdf",
            type="primary"
        )

else:
    st.info("Pastikan file Excel data warga RT 06 sudah tersedia.")
