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
from PIL import Image

st.set_page_config(
    page_title="Portal Resmi RT 06 RW 14 Griya Permata Raya",
    layout="wide",
    page_icon="🏠"
)

# Custom CSS responsif untuk Mobile & Desktop, Gradasi Profesional & Tombol Jelas
st.markdown("""
<style>
    .stApp {
        background: linear-gradient(135deg, #f1f5f9 0%, #cbd5e1 100%) !important;
    }
    .main .block-container {
        background: transparent !important;
    }
    .metric-card {
        background: rgba(255, 255, 255, 0.95);
        backdrop-filter: blur(12px);
        border: 1px solid rgba(226, 232, 240, 0.9);
        padding: 24px;
        border-radius: 18px;
        box-shadow: 0 10px 25px -5px rgba(0, 0, 0, 0.08);
        margin-bottom: 12px;
    }
    .metric-title {
        font-size: 13px !important;
        text-transform: uppercase;
        font-weight: 700;
        color: #475569;
        letter-spacing: 0.5px;
    }
    .metric-value {
        font-size: 32px !important;
        font-weight: 900 !important;
        margin-top: 6px;
    }
    .jumbo-title {
        font-size: 44px !important;
        font-weight: 900 !important;
        color: #1e3a8a !important;
        line-height: 1.1 !important;
        margin-bottom: 5px !important;
    }
    .jumbo-subtitle {
        font-size: 20px !important;
        font-weight: 700 !important;
        color: #475569 !important;
    }
    .stButton button {
        font-size: 17px !important;
        font-weight: 800 !important;
        padding: 18px 22px !important;
        border-radius: 14px !important;
        border: none !important;
        color: #ffffff !important;
        text-shadow: 0 1px 2px rgba(0,0,0,0.2);
        box-shadow: 0 8px 15px rgba(0,0,0,0.1) !important;
        transition: all 0.3s ease !important;
        width: 100% !important;
    }
    .stButton button:hover {
        transform: translateY(-3px) !important;
        box-shadow: 0 12px 20px rgba(0,0,0,0.15) !important;
        color: #ffffff !important;
    }
    div.stButton:nth-of-type(1) button { background: linear-gradient(135deg, #2563eb, #1d4ed8) !important; }
    div.stButton:nth-of-type(2) button { background: linear-gradient(135deg, #059669, #047857) !important; }
    div.stButton:nth-of-type(3) button { background: linear-gradient(135deg, #0284c7, #0369a1) !important; }
    div.stButton:nth-of-type(4) button { background: linear-gradient(135deg, #db2777, #be185d) !important; }
    div.stButton:nth-of-type(5) button { background: linear-gradient(135deg, #d97706, #b45309) !important; }
    div.stButton:nth-of-type(6) button { background: linear-gradient(135deg, #7c3aed, #6d28d9) !important; }
    div.stButton:nth-of-type(7) button { background: linear-gradient(135deg, #ea580c, #c2410c) !important; }
</style>
""", unsafe_allow_html=True)

if 'selected_menu' not in st.session_state:
    st.session_state.selected_menu = "Beranda / Dashboard"

# Inisialisasi State Data Kas RT Awal (Tanpa kolom No)
if 'df_kas_rt_state' not in st.session_state:
    st.session_state.df_kas_rt_state = pd.DataFrame({
        "Tanggal": ["01/06/2026", "05/06/2026", "12/06/2026", "20/06/2026"],
        "Uraian / Keterangan Transaksi": [
            "Saldo Awal Periode Lalu", 
            "Penerimaan Iuran Warga Bulanan (Periode Juni)", 
            "Pengeluaran Perbaikan Lampu Penerangan Jalan RT", 
            "Pengeluaran Konsumsi Rapat Koordinasi Warga"
        ],
        "Debet (Masuk)": ["1500000", "2400000", "", ""],
        "Kredit (Keluar)": ["", "", "350000", "150000"]
    })

# Inisialisasi State Data Kas Sosial Awal (Tanpa kolom No, dimulai langsung dari Tanggal)
if 'df_kas_sosial_state' not in st.session_state:
    st.session_state.df_kas_sosial_state = pd.DataFrame({
        "Tanggal": ["01/06/2026", "05/06/2026", "10/06/2026", "15/06/2026", "20/06/2026", "25/06/2026", "28/06/2026", "30/06/2026"],
        "Uraian / Keterangan Transaksi": [
            "Saldo Awal Kotak Sosial Perelek",
            "Penerimaan Perelek Warga Minggu ke-1",
            "Penerimaan Perelek Warga Minggu ke-2",
            "Pengeluaran Bantuan Warga Sakit",
            "Penerimaan Perelek Warga Minggu ke-3",
            "Pengeluaran Santunan Kedukaan",
            "Penerimaan Perelek Warga Minggu ke-4",
            "Saldo Akhir Kas Sosial"
        ],
        "Debet (Masuk)": ["750000", "150000", "150000", "", "150000", "", "150000", ""],
        "Kredit (Keluar)": ["", "", "", "200000", "", "250000", "", ""]
    })

def parsing_angka_aman(val):
    if pd.isna(val) or val == "" or val == "-":
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    val_str = str(val).strip().replace("Rp", "").replace(" ", "")
    clean_str = "".join([c for c in val_str if c.isdigit() or c == '-' or c == '.'])
    if '.' in clean_str and ',' in clean_str:
        clean_str = clean_str.replace('.', '').replace(',', '.')
    elif clean_str.count('.') > 1:
        clean_str = clean_str.replace('.', '')
    try:
        return float(clean_str) if clean_str != "" else 0.0
    except:
        return 0.0

def format_Rupiah(num):
    try:
        n = float(num)
        if n == 0: return "-"
        return f"Rp {int(n):,}".replace(",", ".")
    except:
        return "-"

def hitung_dan_tampilkan_tabel_tunggal(df_input):
    df = df_input.copy()
    saldo_list = []
    curr = 0.0
    
    saldo_formatted = []
    
    for idx, row in df.iterrows():
        deb = parsing_angka_aman(row.get("Debet (Masuk)", 0))
        kre = parsing_angka_aman(row.get("Kredit (Keluar)", 0))
        
        if idx == 0:
            curr = deb - kre
        else:
            curr = curr + deb - kre
            
        saldo_list.append(curr)
        saldo_formatted.append(format_Rupiah(curr))
        
    df["Saldo (Rp)"] = saldo_formatted
    return df

# Header Utama Portal RT 06
col_logo, col_title = st.columns([1, 3.5])
with col_logo:
    logo_path = "logo_rt06.png"
    if not os.path.exists(logo_path):
        logo_path = "logo_rt06.jpg"
    
    if os.path.exists(logo_path):
        try:
            img = Image.open(logo_path).convert("RGBA")
            datas = img.getdata()
            new_data = []
            for item in datas:
                if item[0] > 240 and item[1] > 240 and item[2] > 240:
                    new_data.append((255, 255, 255, 0))
                else:
                    new_data.append(item)
            img.putdata(new_data)
            st.image(img, width=380)
        except:
            st.image(logo_path, width=380)
    else:
        st.write("🏠")

with col_title:
    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown('<div class="jumbo-title">🏠 PORTAL RT 06 / RW 14</div>', unsafe_allow_html=True)
    st.markdown('<div class="jumbo-subtitle">Griya Permata Raya • Desa Nanjung Mekar, Rancaekek</div>', unsafe_allow_html=True)

st.write("---")

@st.cache_data(ttl=60)
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
        
        for col in df.select_dtypes(include=['object']).columns:
            df[col] = df[col].astype(str).str.strip()
            df.loc[df[col].str.lower() == 'nan', col] = None
        
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
    col_kk_candi = [c for c in df.columns if "KEPALA" in c or "KK" in c]
    col_kk = col_kk_candi[0] if col_kk_candi else df.columns[2]
    
    col_nama_candi = [c for c in df.columns if "ANGGOTA" in c or "NAMA" in c]
    col_nama = col_nama_candi[0] if col_nama_candi else df.columns[3]
    
    col_rumah_candi = [c for c in df.columns if "RUMAH" in c or "ALAMAT" in c]
    col_rumah = col_rumah_candi[0] if col_rumah_candi else None

    col_jk = next((c for c in df.columns if "JK" in c or "KELAMIN" in c or "GENDER" in c), None)
    col_usia = next((c for c in df.columns if "USIA" in c or "UMUR" in c), None)

    st.sidebar.markdown("### 🧭 Navigasi Menu")
    daftar_menu_pilihan = [
        "Beranda / Dashboard",
        "📋 Data Seluruh Warga",
        "🗂️ Cetak Kartu Keluarga (KK)", 
        "📈 Grafik Demografi", 
        "🛠️ Kelola Warga", 
        "📊 Rekapitulasi Administrasi RW",
        "💰 Laporan Kas RT & Sosial (Perelek R6 Suayunan)",
        "🖨️ Cetak Rekap PDF"
    ]
    
    selected_sidebar = st.sidebar.selectbox("Pilih Halaman:", daftar_menu_pilihan, index=daftar_menu_pilihan.index(st.session_state.selected_menu) if st.session_state.selected_menu in daftar_menu_pilihan else 0)

    if selected_sidebar != st.session_state.selected_menu:
        st.session_state.selected_menu = selected_sidebar
        st.rerun()

    menu = st.session_state.selected_menu

    if menu == "Beranda / Dashboard":
        col_jam1, col_jam2 = st.columns([2, 2])
        with col_jam1:
            st.subheader("📊 Dashboard Eksekutif")
        with col_jam2:
            placeholder_waktu = st.empty()

        df_ffill = df.copy()
        if col_kk:
            df_ffill[col_kk] = df_ffill[col_kk].replace('', pd.NA).ffill()
        if col_rumah:
            df_ffill[col_rumah] = df_ffill[col_rumah].replace('', pd.NA).ffill()

        total_jiwa = len(df)
        total_kk = df_ffill[col_kk].nunique() if col_kk in df_ffill.columns else 0
        
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
        <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr)); gap: 15px; margin-top: 10px; margin-bottom: 25px;">
            <div class="metric-card" style="border-left: 6px solid #2563eb;">
                <div class="metric-title">Jumlah KK</div>
                <div class="metric-value" style="color: #1e3a8a;">{total_kk} <span style="font-size: 16px; font-weight: 600; color: #64748b;">KK</span></div>
            </div>
            <div class="metric-card" style="border-left: 6px solid #059669;">
                <div class="metric-title">Total Jiwa</div>
                <div class="metric-value" style="color: #065f46;">{total_jiwa} <span style="font-size: 16px; font-weight: 600; color: #64748b;">Jiwa</span></div>
            </div>
            <div class="metric-card" style="border-left: 6px solid #0284c7;">
                <div class="metric-title">Laki-laki</div>
                <div class="metric-value" style="color: #0369a1;">{jml_l} <span style="font-size: 16px; font-weight: 600; color: #64748b;">Orang</span></div>
            </div>
            <div class="metric-card" style="border-left: 6px solid #db2777;">
                <div class="metric-title">Perempuan</div>
                <div class="metric-value" style="color: #9d174d;">{jml_p} <span style="font-size: 16px; font-weight: 600; color: #64748b;">Orang</span></div>
            </div>
            <div class="metric-card" style="border-left: 6px solid #d97706;">
                <div class="metric-title">Balita (0-5 th)</div>
                <div class="metric-value" style="color: #b45309;">{jml_balita} <span style="font-size: 16px; font-weight: 600; color: #64748b;">Jiwa</span></div>
            </div>
            <div class="metric-card" style="border-left: 6px solid #7c3aed;">
                <div class="metric-title">Lansia (>60 th)</div>
                <div class="metric-value" style="color: #5b21b6;">{jml_lansia} <span style="font-size: 16px; font-weight: 600; color: #64748b;">Jiwa</span></div>
            </div>
        </div>
        """, unsafe_allow_html=True)

        st.write("---")
        st.markdown("### 🚀 Menu Utama Portal RT 06")
        
        col_m1, col_m2 = st.columns(2)
        with col_m1:
            if st.button("📋 Data Seluruh Warga", use_container_width=True):
                st.session_state.selected_menu = "📋 Data Seluruh Warga"
                st.rerun()
            if st.button("🗂️ Cetak Kartu Keluarga (KK)", use_container_width=True):
                st.session_state.selected_menu = "🗂️ Cetak Kartu Keluarga (KK)"
                st.rerun()
            if st.button("📊 Rekapitulasi Administrasi RW", use_container_width=True):
                st.session_state.selected_menu = "📊 Rekapitulasi Administrasi RW"
                st.rerun()
            if st.button("💰 Laporan Kas RT & Sosial (Perelek R6)", use_container_width=True):
                st.session_state.selected_menu = "💰 Laporan Kas RT & Sosial (Perelek R6 Suayunan)"
                st.rerun()
        with col_m2:
            if st.button("📈 Grafik Demografi", use_container_width=True):
                st.session_state.selected_menu = "📈 Grafik Demografi"
                st.rerun()
            if st.button("🛠️ Kelola Data Warga", use_container_width=True):
                st.session_state.selected_menu = "🛠️ Kelola Warga"
                st.rerun()
            if st.button("🖨️ Cetak Laporan Rekap PDF", use_container_width=True):
                st.session_state.selected_menu = "🖨️ Cetak Rekap PDF"
                st.rerun()

        for _ in range(5):
            waktu_sekarang = datetime.now().strftime("%d %B %Y | %H:%M:%S")
            placeholder_waktu.markdown(f"""
            <div style="background: rgba(255, 255, 255, 0.9); border: 1px solid #cbd5e1; padding: 8px 12px; border-radius: 10px; text-align: right;">
                <span style="font-size: 10px; color: #64748b;">🕒 Live Update:</span><br>
                <strong style="font-size: 12px; color: #0f172a;">{waktu_sekarang}</strong>
            </div>
            """, unsafe_allow_html=True)
            time.sleep(1)
        st.rerun()

    elif menu == "📋 Data Seluruh Warga":
        if st.button("⬅️ Kembali ke Beranda"):
            st.session_state.selected_menu = "Beranda / Dashboard"
            st.rerun()
        st.subheader("📋 Data Keseluruhan Warga")
        st.data_editor(df, num_rows="dynamic", use_container_width=True, key="editor_warga_grid")

    elif menu == "🗂️ Cetak Kartu Keluarga (KK)":
        if st.button("⬅️ Kembali ke Beranda"):
            st.session_state.selected_menu = "Beranda / Dashboard"
            st.rerun()
        st.subheader("🗂️ Cetak Kartu Keluarga (KK)")
        
        df_ffill = df.copy()
        if col_kk:
            df_ffill[col_kk] = df_ffill[col_kk].replace('', pd.NA).ffill()
        if col_rumah:
            df_ffill[col_rumah] = df_ffill[col_rumah].replace('', pd.NA).ffill()

        daftar_kk = df_ffill[col_kk].dropna().astype(str).str.strip()
        daftar_kk = sorted(list(set([x for x in daftar_kk if x != "" and x.lower() != "nan" and x.lower() != "none"])))
        
        pilihan_kk = st.selectbox("Pilih Kepala Keluarga:", daftar_kk)
        
        if pilihan_kk:
            df_keluarga = df_ffill[df_ffill[col_kk].astype(str).str.strip().str.lower() == pilihan_kk.strip().lower()].copy()
            no_rmh = str(df_keluarga[col_rumah].dropna().iloc[0]) if col_rumah and not df_keluarga[df_keluarga[col_rumah].notna()].empty else "-"
            
            cols_tampilan_web = [c for c in df_keluarga.columns if c != col_kk and "URUT" not in c and c != "NO"]
            
            st.markdown(f"""
            <div style="background: #ffffff; border: 2px solid #2563eb; border-radius: 12px; padding: 12px; margin-bottom: 12px;">
                <h4 style="margin: 0; color: #1e3a8a; font-size: 15px;">🏠 No. Rumah: {no_rmh} | KK: {pilihan_kk}</h4>
                <p style="margin: 4px 0 0 0; font-size: 13px; color: #64748b;">Jumlah Anggota Keluarga Terdata: <b>{len(df_keluarga)} Jiwa</b></p>
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
                label=f"📥 Download PDF KK ({pilihan_kk})",
                data=pdf_kk_bytes,
                file_name=f"KK_{pilihan_kk.replace(' ', '_')}.pdf",
                mime="application/pdf",
                type="primary"
            )

    elif menu == "📈 Grafik Demografi":
        if st.button("⬅️ Kembali ke Beranda"):
            st.session_state.selected_menu = "Beranda / Dashboard"
            st.rerun()
        st.subheader("📈 Analisis & Statistik Grafik Demografi Warga")
        
        col_pend = next((c for c in df.columns if "PENDIDIKAN" in c), None)
        col_pek = next((c for c in df.columns if "PEKERJAAN" in c), None)
        col_status = next((c for c in df.columns if "STATUS" in c and "KAWIN" in c) or (c for c in df.columns if "STATUS" in c), None)
        
        chart_font = dict(size=15, family="Arial, sans-serif")
        title_font = dict(size=20, family="Arial, sans-serif")

        if col_jk:
            df_jk = df[col_jk].dropna().value_counts().reset_index()
            df_jk.columns = ["Jenis Kelamin", "Jumlah"]
            fig_jk = px.pie(df_jk, names="Jenis Kelamin", values="Jumlah", hole=0.5, title="👥 Rasio Penduduk Berdasarkan Jenis Kelamin", color_discrete_sequence=px.colors.qualitative.Bold)
            fig_jk.update_traces(textfont_size=18, textinfo="percent+label+value")
            fig_jk.update_layout(font=chart_font, title_font=title_font, legend=dict(font=dict(size=14)))
            st.plotly_chart(fig_jk, use_container_width=True)

        if col_status:
            st.markdown("---")
            df_st = df[col_status].dropna().value_counts().reset_index()
            df_st.columns = ["Status", "Jumlah"]
            fig_st = px.bar(df_st, x="Status", y="Jumlah", text="Jumlah", title="💍 Distribusi Status Pernikahan Warga", color="Status", color_discrete_sequence=px.colors.qualitative.Pastel)
            fig_st.update_traces(textfont_size=16, textposition="outside")
            fig_st.update_layout(font=chart_font, title_font=title_font)
            st.plotly_chart(fig_st, use_container_width=True)

        if col_pend:
            st.markdown("---")
            df_pd = df[col_pend].dropna().value_counts().reset_index()
            df_pd.columns = ["Pendidikan", "Jumlah"]
            fig_pd = px.bar(df_pd, x="Pendidikan", y="Jumlah", text="Jumlah", title="🎓 Tingkat Pendidikan Terakhir Warga", color="Pendidikan", color_discrete_sequence=px.colors.qualitative.Vivid)
            fig_pd.update_traces(textfont_size=16, textposition="outside")
            fig_pd.update_layout(font=chart_font, title_font=title_font)
            st.plotly_chart(fig_pd, use_container_width=True)

        if col_pek:
            st.markdown("---")
            df_pk = df[col_pek].dropna().value_counts().reset_index()
            df_pk.columns = ["Pekerjaan", "Jumlah"]
            fig_pk = px.bar(df_pk, x="Pekerjaan", y="Jumlah", text="Jumlah", title="💼 Distribusi Mata Pencaharian / Pekerjaan Warga", color="Pekerjaan", color_discrete_sequence=px.colors.qualitative.Safe)
            fig_pk.update_traces(textfont_size=16, textposition="outside")
            fig_pk.update_layout(font=chart_font, title_font=title_font)
            st.plotly_chart(fig_pk, use_container_width=True)

        if col_usia:
            st.markdown("---")
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
            
            fig_usia = px.pie(df_usia_count, names="Kategori Usia", values="Jumlah", hole=0.4, title="👶 Kelompok Rentang Usia Penduduk", color_discrete_sequence=px.colors.qualitative.Set3)
            fig_usia.update_traces(textfont_size=18, textinfo="percent+label+value")
            fig_usia.update_layout(font=chart_font, title_font=title_font)
            st.plotly_chart(fig_usia, use_container_width=True)

    elif menu == "🛠️ Kelola Warga":
        if st.button("⬅️ Kembali ke Beranda"):
            st.session_state.selected_menu = "Beranda / Dashboard"
            st.rerun()
        st.subheader("🛠️ Kelola Data Warga")
        aksi = st.selectbox("Aksi:", ["➕ Tambah Warga Baru"])
        
        if aksi == "➕ Tambah Warga Baru":
            with st.form("form_tambah"):
                df_ffill_kelola = df.copy()
                if col_kk:
                    df_ffill_kelola[col_kk] = df_ffill_kelola[col_kk].replace('', pd.NA).ffill()
                if col_rumah:
                    df_ffill_kelola[col_rumah] = df_ffill_kelola[col_rumah].replace('', pd.NA).ffill()

                daftar_no_rumah = sorted(list(set(df_ffill_kelola[col_rumah].dropna().astype(str).tolist()))) if col_rumah else ["B3-01", "B3-02"]
                no_rumah = st.selectbox("No. Rumah", daftar_no_rumah)
                
                nama_kk = st.selectbox("Nama Kepala Keluarga", sorted(list(set(df_ffill_kelola[col_kk].dropna().astype(str).tolist()))))
                nama_anggota = st.text_input("Nama Lengkap Anggota Keluarga")
                jk = st.selectbox("Jenis Kelamin", ["L", "P"])
                hubungan = st.selectbox("Hubungan Keluarga", ["Kepala Keluarga", "Istri", "Anak Kandung", "Famili Lain", "Mertua"])
                tempat_lahir = st.text_input("Tempat Lahir")
                
                tanggal_lahir_date = st.date_input("Tanggal Lahir", value=date(1995, 1, 1))
                bulan_indo_nama = {1: "Januari", 2: "Februari", 3: "Maret", 4: "April", 5: "Mei", 6: "Juni", 7: "Juli", 8: "Agustus", 9: "September", 10: "Oktober", 11: "November", 12: "Desember"}
                tanggal_lahir_str = f"{tanggal_lahir_date.day:02d} {bulan_indo_nama.get(tanggal_lahir_date.month, '')} {tanggal_lahir_date.year}"
                
                usia = st.number_input("Usia", min_value=0, max_value=120, value=25)
                status_nikah = st.selectbox("Status Pernikahan", ["Belum Kawin", "Kawin", "Cerai Hidup", "Cerai Mati"])
                
                pendidikan = st.selectbox("Pendidikan Terakhir", [
                    "Tamat SLTA/sederajat", "Tamat SLTP/sederajat", "Tamat SD/sederajat", 
                    "Diploma / Sarjana (S1/S2/S3)", "Belum / Tidak Sekolah"
                ])
                
                pekerjaan = st.selectbox("Pekerjaan", [
                    "Karyawan Swasta", "Wiraswasta", "Mengurus Rumah Tangga", 
                    "Belum Bekerja", "Pelajar / Mahasiswa", "PNS / TNI / Polri"
                ])
                
                status_rumah = st.selectbox("Status Rumah", ["Milik / Tetap", "Sewa / Kontrak"])
                status_domisili = st.selectbox("Status Domisili", ["Warga Tetap", "Warga Kontrak", "Luar NM"])
                
                if st.form_submit_button("Simpan Data Warga Baru"):
                    st.success(f"Data **{nama_anggota}** berhasil disiapkan!")

    elif menu == "📊 Rekapitulasi Administrasi RW":
        if st.button("⬅️ Kembali ke Beranda"):
            st.session_state.selected_menu = "Beranda / Dashboard"
            st.rerun()
        st.subheader("📊 Rekapitulasi Administrasi RW")
        
        df_ffill_rw = df.copy()
        if col_kk:
            df_ffill_rw[col_kk] = df_ffill_rw[col_kk].replace('', pd.NA).ffill()
        if col_rumah:
            df_ffill_rw[col_rumah] = df_ffill_rw[col_rumah].replace('', pd.NA).ffill()

        total_kk_rw = df_ffill_rw[col_kk].nunique() if col_kk in df_ffill_rw.columns else 0
        total_jiwa_rw = len(df)
        jml_l_rw = len(df[df[col_jk].astype(str).str.upper().str.contains("L")]) if col_jk else 0
        jml_p_rw = len(df[df[col_jk].astype(str).str.upper().str.contains("P")]) if col_jk else 0
        
        balita_rw, lansia_rw = 0, 0
        if col_usia:
            u_ser = df[col_usia].apply(lambda x: int(x) if str(x).isdigit() else -1)
            balita_rw = len(u_ser[(u_ser >= 0) & (u_ser <= 5)])
            lansia_rw = len(u_ser[u_ser > 60])

        data_rekap_rw = {
            "No": [1, 2, 3, 4, 5, 6, 7],
            "Komponen Rekapitulasi Administrasi": [
                "Jumlah Kepala Keluarga (KK)",
                "Jumlah Jiwa / Penduduk Total",
                "Jumlah Penduduk Laki-Laki",
                "Jumlah Penduduk Perempuan",
                "Jumlah Balita (0-5 Tahun)",
                "Jumlah Lansia (>60 Tahun)",
                "Jumlah Total Rumah Terdata"
            ],
            "Jumlah / Volume": [
                f"{total_kk_rw} KK",
                f"{total_jiwa_rw} Jiwa",
                f"{jml_l_rw} Orang",
                f"{jml_p_rw} Orang",
                f"{balita_rw} Jiwa",
                f"{lansia_rw} Jiwa",
                f"{df_ffill_rw[col_rumah].nunique() if col_rumah in df_ffill_rw.columns else 0} Rumah"
            ]
        }
        df_rekap_rw = pd.DataFrame(data_rekap_rw)
        st.dataframe(df_rekap_rw, use_container_width=True, hide_index=True)

    elif menu == "💰 Laporan Kas RT & Sosial (Perelek R6 Suayunan)":
        if st.button("⬅️ Kembali ke Beranda"):
            st.session_state.selected_menu = "Beranda / Dashboard"
            st.rerun()
        st.subheader("💰 Laporan Keuangan Kas RT & Kas Sosial (Perelek R6 Suayunan)")
        st.markdown("💡 **Info Super Stabil:** Semua kolom dikunci murni sebagai teks (`TextColumn`) tanpa kolom nomor urut. Data yang Anda salin dari Excel akan masuk dengan akurat, tidak berubah sendiri, dan kolom Saldo otomatis menghitung secara benar.")
        
        def format_rupiah_pdf(num):
            try:
                n = float(num)
                if n == 0: return "-"
                return f"Rp {int(n):,}".replace(",", ".")
            except:
                return "-"

        tab_kas1, tab_kas2 = st.tabs(["📊 Buku Kas RT 06", "🌾 Buku Kas Sosial (Perelek)"])
        
        with tab_kas1:
            st.markdown("### Buku Kas RT 06")
            
            df_rt_view = st.session_state.df_kas_rt_state.copy()
            df_rt_view_saldo = hitung_dan_tampilkan_tabel_tunggal(df_rt_view)
            
            edited_rt = st.data_editor(
                df_rt_view_saldo, 
                num_rows="dynamic", 
                use_container_width=True, 
                key="editor_kas_rt_pure_text",
                column_config={
                    "Tanggal": st.column_config.TextColumn("Tanggal"),
                    "Uraian / Keterangan Transaksi": st.column_config.TextColumn("Uraian / Keterangan Transaksi"),
                    "Debet (Masuk)": st.column_config.TextColumn("Debet (Masuk)"),
                    "Kredit (Keluar)": st.column_config.TextColumn("Kredit (Keluar)"),
                    "Saldo (Rp)": st.column_config.TextColumn("Saldo (Rp)", disabled=True)
                }
            )
            
            if not edited_rt.empty:
                st.session_state.df_kas_rt_state = pd.DataFrame({
                    "Tanggal": edited_rt.get("Tanggal", "").astype(str),
                    "Uraian / Keterangan Transaksi": edited_rt.get("Uraian / Keterangan Transaksi", "").astype(str),
                    "Debet (Masuk)": edited_rt.get("Debet (Masuk)", "").astype(str),
                    "Kredit (Keluar)": edited_rt.get("Kredit (Keluar)", "").astype(str)
                })

            def buat_pdf_standar_akuntansi(df_lap, judul):
                buffer = io.BytesIO()
                doc = SimpleDocTemplate(buffer, pagesize=letter, rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=30)
                elements = []
                styles = getSampleStyleSheet()
                
                elements.append(Paragraph("PEMERINTAH KABUPATEN BANDUNG", ParagraphStyle('Sub1', parent=styles['Normal'], alignment=1, fontSize=10, textColor=colors.gray)))
                elements.append(Paragraph("RT 06 / RW 14 - KECAMATAN RANCAAEKEK", ParagraphStyle('Sub2', parent=styles['Normal'], alignment=1, fontSize=10, textColor=colors.gray)))
                elements.append(Paragraph(judul, ParagraphStyle('Title', parent=styles['Heading1'], fontSize=13, alignment=1, textColor=colors.HexColor('#1f2937'))))
                elements.append(Spacer(1, 15))
                
                df_pdf_clean = hitung_dan_tampilkan_tabel_tunggal(df_lap)
                df_pdf_clean["Debet (Masuk)"] = df_pdf_clean["Debet (Masuk)"].apply(parsing_angka_aman).apply(format_rupiah_pdf)
                df_pdf_clean["Kredit (Keluar)"] = df_pdf_clean["Kredit (Keluar)"].apply(parsing_angka_aman).apply(format_rupiah_pdf)
                
                kolom = list(df_pdf_clean.columns)
                cell_s = ParagraphStyle('Cell', parent=styles['Normal'], fontSize=8.5, leading=10, alignment=1)
                head_s = ParagraphStyle('Head', parent=styles['Normal'], fontSize=9, leading=11, textColor=colors.whitesmoke, fontName='Helvetica-Bold', alignment=1)
                
                t_data = [[Paragraph(c, head_s) for c in kolom]]
                for _, r in df_pdf_clean.iterrows():
                    t_data.append([Paragraph(str(r[c]), cell_s) for c in kolom])
                    
                t = Table(t_data, colWidths=[70, 230, 75, 75, 75], repeatRows=1)
                t.setStyle(TableStyle([
                    ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#1e3a8a')),
                    ('ALIGN', (0,0), (-1,-1), 'CENTER'),
                    ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                    ('BOTTOMPADDING', (0,0), (-1,-1), 5),
                    ('TOPPADDING', (0,0), (-1,-1), 5),
                    ('BACKGROUND', (0,1), (-1,-1), colors.HexColor('#f8fafc')),
                    ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#cbd5e1')),
                ]))
                elements.append(t)
                elements.append(Spacer(1, 20))
                
                ttd_data = [
                    [Paragraph("<b>Mengetahui,<br/>Ketua RT 06</b>", ParagraphStyle('T1', parent=styles['Normal'], alignment=1, fontSize=9)),
                     Paragraph("<b>Bendahara RT 06</b>", ParagraphStyle('T2', parent=styles['Normal'], alignment=1, fontSize=9))],
                    [Spacer(1, 35), Spacer(1, 35)],
                    [Paragraph("<b>( ......................................... )</b>", ParagraphStyle('T3', parent=styles['Normal'], alignment=1, fontSize=9)),
                     Paragraph("<b>( ......................................... )</b>", ParagraphStyle('T4', parent=styles['Normal'], alignment=1, fontSize=9))]
                ]
                t_ttd = Table(ttd_data, colWidths=[250, 250])
                elements.append(t_ttd)
                
                doc.build(elements)
                buffer.seek(0)
                return buffer.getvalue()

            pdf_akuntansi_rt = buat_pdf_standar_akuntansi(st.session_state.df_kas_rt_state, "LAPORAN PERTANGGUNGJAWABAN KEUANGAN KAS RT 06")
            st.download_button(
                label="📥 Download PDF Laporan Standar Akuntansi Kas RT",
                data=pdf_akuntansi_rt,
                file_name="Laporan_Akuntansi_Kas_RT06.pdf",
                mime="application/pdf",
                type="primary"
            )

        with tab_kas2:
            st.markdown("### Buku Kas Sosial / Perelek")
            
            df_sosial_view = st.session_state.df_kas_sosial_state.copy()
            df_sosial_view_saldo = hitung_dan_tampilkan_tabel_tunggal(df_sosial_view)
            
            edited_sosial = st.data_editor(
                df_sosial_view_saldo, 
                num_rows="dynamic", 
                use_container_width=True, 
                key="editor_kas_sosial_pure_text",
                column_config={
                    "Tanggal": st.column_config.TextColumn("Tanggal"),
                    "Uraian / Keterangan Transaksi": st.column_config.TextColumn("Uraian / Keterangan Transaksi"),
                    "Debet (Masuk)": st.column_config.TextColumn("Debet (Masuk)"),
                    "Kredit (Keluar)": st.column_config.TextColumn("Kredit (Keluar)"),
                    "Saldo (Rp)": st.column_config.TextColumn("Saldo (Rp)", disabled=True)
                }
            )
            
            if not edited_sosial.empty:
                st.session_state.df_kas_sosial_state = pd.DataFrame({
                    "Tanggal": edited_sosial.get("Tanggal", "").astype(str),
                    "Uraian / Keterangan Transaksi": edited_sosial.get("Uraian / Keterangan Transaksi", "").astype(str),
                    "Debet (Masuk)": edited_sosial.get("Debet (Masuk)", "").astype(str),
                    "Kredit (Keluar)": edited_sosial.get("Kredit (Keluar)", "").astype(str)
                })
            
            pdf_akuntansi_perelek = buat_pdf_standar_akuntansi(st.session_state.df_kas_sosial_state, "LAPORAN DANA SOSIAL PERELEK R6 SUAYUNAN RT 06")
            st.download_button(
                label="📥 Download PDF Laporan Standar Akuntansi Kas Sosial",
                data=pdf_akuntansi_perelek,
                file_name="Laporan_Akuntansi_Kas_Sosial.pdf",
                mime="application/pdf",
                type="primary"
            )

    elif menu == "🖨️ Cetak Rekap PDF":
        if st.button("⬅️ Kembali ke Beranda"):
            st.session_state.selected_menu = "Beranda / Dashboard"
            st.rerun()
        st.subheader("🖨️ Cetak Rekapitulasi Keseluruhan (PDF)")
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
            label="📥 Download PDF Rekapitulasi",
            data=pdf_rekap,
            file_name="Rekap_Warga_RT06.pdf",
            mime="application/pdf",
            type="primary"
        )

else:
    st.info("Pastikan file Excel data warga RT 06 sudah tersedia.")
