import streamlit as st
import pandas as pd
import plotly.express as px
import os
import io
import time
from datetime import datetime
from zoneinfo import ZoneInfo
from reportlab.lib.pagesizes import letter, landscape
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image as RLImage
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from PIL import Image

st.set_page_config(
    page_title="Sistem Administrasi Kependudukan (SIAK) - RT 06 / RW 14",
    layout="wide",
    page_icon="🏛️"
)

# Custom CSS Profesional - Menjamin Background Stabil & Teks Tombol Terlihat Jelas
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
        font-size: 40px !important;
        font-weight: 900 !important;
        color: #1e3a8a !important;
        line-height: 1.1 !important;
        margin-bottom: 5px !important;
    }
    .jumbo-subtitle {
        font-size: 18px !important;
        font-weight: 700 !important;
        color: #475569 !important;
    }
    .stButton button {
        font-size: 16px !important;
        font-weight: 800 !important;
        padding: 18px 22px !important;
        border-radius: 14px !important;
        border: none !important;
        color: #ffffff !important;
        -webkit-text-fill-color: #ffffff !important;
        text-shadow: 0 1px 2px rgba(0,0,0,0.3);
        box-shadow: 0 8px 15px rgba(0,0,0,0.1) !important;
        transition: all 0.3s ease !important;
        width: 100% !important;
    }
    .stButton button:hover {
        transform: translateY(-3px) !important;
        box-shadow: 0 12px 20px rgba(0,0,0,0.15) !important;
        color: #ffffff !important;
        -webkit-text-fill-color: #ffffff !important;
    }
    div.stButton:nth-of-type(1) button { background: linear-gradient(135deg, #2563eb, #1d4ed8) !important; }
    div.stButton:nth-of-type(2) button { background: linear-gradient(135deg, #0d9488, #0f766e) !important; }
    div.stButton:nth-of-type(3) button { background: linear-gradient(135deg, #059669, #047857) !important; }
    div.stButton:nth-of-type(4) button { background: linear-gradient(135deg, #0284c7, #0369a1) !important; }
    div.stButton:nth-of-type(5) button { background: linear-gradient(135deg, #db2777, #be185d) !important; }
    div.stButton:nth-of-type(6) button { background: linear-gradient(135deg, #d97706, #b45309) !important; }
    div.stButton:nth-of-type(7) button { background: linear-gradient(135deg, #7c3aed, #6d28d9) !important; }
    div.stButton:nth-of-type(8) button { background: linear-gradient(135deg, #ea580c, #c2410c) !important; }
</style>
""", unsafe_allow_html=True)

if 'selected_menu' not in st.session_state:
    st.session_state.selected_menu = "Dashboard Eksekutif Kependudukan"

FILE_KAS_RT = "penyimpanan_kas_rt.csv"
FILE_KAS_SOSIAL = "penyimpanan_kas_sosial.csv"
FILE_EXCEL_WARGA = "data_warga_rt06.xlsx"

def muat_data_kas(file_path, default_df):
    if os.path.exists(file_path):
        try:
            df_disk = pd.read_csv(file_path)
            if not df_disk.empty:
                df_disk['_dt_sort'] = pd.to_datetime(df_disk['Tanggal'], format='%d/%m/%Y', errors='coerce')
                df_disk = df_disk.sort_values(by='_dt_sort').drop(columns=['_dt_sort']).reset_index(drop=True)
                return df_disk
        except:
            pass
    if not default_df.empty:
        default_df['_dt_sort'] = pd.to_datetime(default_df['Tanggal'], format='%d/%m/%Y', errors='coerce')
        default_df = default_df.sort_values(by='_dt_sort').drop(columns=['_dt_sort']).reset_index(drop=True)
    return default_df

def simpan_data_kas(file_path, df):
    try:
        df_s = df.copy()
        if not df_s.empty:
            df_s['_dt_sort'] = pd.to_datetime(df_s['Tanggal'], format='%d/%m/%Y', errors='coerce')
            df_s = df_s.sort_values(by='_dt_sort').drop(columns=['_dt_sort']).reset_index(drop=True)
        df_s.to_csv(file_path, index=False)
    except:
        pass

if 'df_kas_rt_state' not in st.session_state:
    default_rt = pd.DataFrame(columns=["Tanggal", "Uraian / Keterangan Transaksi", "Debet (Masuk)", "Kredit (Keluar)"])
    st.session_state.df_kas_rt_state = muat_data_kas(FILE_KAS_RT, default_rt)

if 'df_kas_sosial_state' not in st.session_state:
    default_sosial = pd.DataFrame([
        {"Tanggal": "01/01/2026", "Uraian / Keterangan Transaksi": "Saldo Kas Perelek Akhir Desember 2025", "Debet (Masuk)": 350000.0, "Kredit (Keluar)": 0.0},
        {"Tanggal": "10/01/2026", "Uraian / Keterangan Transaksi": "Menjenguk B3 No.17 (dirawat di RS,operasi )", "Debet (Masuk)": 0.0, "Kredit (Keluar)": 200000.0},
        {"Tanggal": "12/01/2026", "Uraian / Keterangan Transaksi": "Menjenguk B3 No.27 (istrinya sakit, tangannya kena air panas)", "Debet (Masuk)": 0.0, "Kredit (Keluar)": 100000.0},
        {"Tanggal": "25/01/2026", "Uraian / Keterangan Transaksi": "Hasil penarikan perelek ke 48", "Debet (Masuk)": 250000.0, "Kredit (Keluar)": 0.0},
        {"Tanggal": "27/01/2026", "Uraian / Keterangan Transaksi": "Menjenguk B3 No.12 (anaknya dirawat di RS,operasi )", "Debet (Masuk)": 0.0, "Kredit (Keluar)": 200000.0},
        {"Tanggal": "25/02/2026", "Uraian / Keterangan Transaksi": "Hasil penarikan perelek ke 49", "Debet (Masuk)": 300000.0, "Kredit (Keluar)": 0.0},
        {"Tanggal": "07/03/2026", "Uraian / Keterangan Transaksi": "Menjenguk Ust.Burhan (Dirawat di RS, DBD)", "Debet (Masuk)": 0.0, "Kredit (Keluar)": 100000.0},
        {"Tanggal": "07/03/2026", "Uraian / Keterangan Transaksi": "Santunan anak yatim PHBI Nuzulul Quran di masjid UBK", "Debet (Masuk)": 0.0, "Kredit (Keluar)": 200000.0},
        {"Tanggal": "30/03/2026", "Uraian / Keterangan Transaksi": "Hasil penarikan perelek ke 50", "Debet (Masuk)": 300000.0, "Kredit (Keluar)": 0.0},
        {"Tanggal": "26/04/2026", "Uraian / Keterangan Transaksi": "Menjenguk B6 No.08 ( ibunya dirawat di RS,operasi tangan )", "Debet (Masuk)": 0.0, "Kredit (Keluar)": 200000.0},
        {"Tanggal": "27/04/2026", "Uraian / Keterangan Transaksi": "Hasil penarikan perelek ke 51", "Debet (Masuk)": 250000.0, "Kredit (Keluar)": 0.0},
        {"Tanggal": "25/05/2026", "Uraian / Keterangan Transaksi": "Hasil penarikan perelek ke 52", "Debet (Masuk)": 150000.0, "Kredit (Keluar)": 0.0},
        {"Tanggal": "24/06/2026", "Uraian / Keterangan Transaksi": "Menjenguk B5 No.12 (sakit selama 4 hari )", "Debet (Masuk)": 0.0, "Kredit (Keluar)": 100000.0},
        {"Tanggal": "26/06/2026", "Uraian / Keterangan Transaksi": "Menjenguk B3 No.18 (anaknys dirawat di RS, sakit typhus )", "Debet (Masuk)": 0.0, "Kredit (Keluar)": 200000.0},
        {"Tanggal": "27/06/2026", "Uraian / Keterangan Transaksi": "Hasil penarikan perelek ke 53", "Debet (Masuk)": 250000.0, "Kredit (Keluar)": 0.0},
        {"Tanggal": "08/07/2025", "Uraian / Keterangan Transaksi": "Menjenguk B3 No.25 (istrinya sakit selama 3 hari )", "Debet (Masuk)": 0.0, "Kredit (Keluar)": 100000.0},
        {"Tanggal": "08/07/2026", "Uraian / Keterangan Transaksi": "Menjenguk B3 No.17 (sakit selama 3 hari)", "Debet (Masuk)": 0.0, "Kredit (Keluar)": 100000.0},
        {"Tanggal": "09/07/2026", "Uraian / Keterangan Transaksi": "Takziah ke rumah B3 No.21 ( org tua meninggal dunia)", "Debet (Masuk)": 0.0, "Kredit (Keluar)": 200000.0},
        {"Tanggal": "23/07/2026", "Uraian / Keterangan Transaksi": "Menjenguk B5 No.01 (Bpk mertua sakit, akibat jatuh dari motor )", "Debet (Masuk)": 0.0, "Kredit (Keluar)": 100000.0},
        {"Tanggal": "29/07/2026", "Uraian / Keterangan Transaksi": "Hasil penarikan perelek ke 54", "Debet (Masuk)": 250000.0, "Kredit (Keluar)": 0.0},
        {"Tanggal": "02/08/2026", "Uraian / Keterangan Transaksi": "Menjenguk B5 No.11 ( istrinya dirawat di RS,operasi )", "Debet (Masuk)": 0.0, "Kredit (Keluar)": 200000.0},
        {"Tanggal": "16/08/2026", "Uraian / Keterangan Transaksi": "Takziah ke rumah B3 No..14 ( org tua meninggal dunia)", "Debet (Masuk)": 0.0, "Kredit (Keluar)": 200000.0},
        {"Tanggal": "18/08/2026", "Uraian / Keterangan Transaksi": "Menjenguk B5 No.12 (sakit selama 3 hari)", "Debet (Masuk)": 0.0, "Kredit (Keluar)": 100000.0},
        {"Tanggal": "21/08/2026", "Uraian / Keterangan Transaksi": "Menjenguk B3 No.06 (istrinya sakit selama 3 hari)", "Debet (Masuk)": 0.0, "Kredit (Keluar)": 100000.0},
        {"Tanggal": "21/08/2026", "Uraian / Keterangan Transaksi": "Menjenguk B3 No.34 (istrinya sakit selama 3 hari)", "Debet (Masuk)": 0.0, "Kredit (Keluar)": 100000.0},
        {"Tanggal": "24/08/2026", "Uraian / Keterangan Transaksi": "KAS RT 06", "Debet (Masuk)": 400000.0, "Kredit (Keluar)": 0.0},
        {"Tanggal": "25/08/2026", "Uraian / Keterangan Transaksi": "Hasil penarikan perelek ke 55", "Debet (Masuk)": 200000.0, "Kredit (Keluar)": 0.0},
        {"Tanggal": "30/08/2026", "Uraian / Keterangan Transaksi": "Menjenguk B5 No.07 ( istrinya dirawat di RS )", "Debet (Masuk)": 0.0, "Kredit (Keluar)": 100000.0},
        {"Tanggal": "07/09/2026", "Uraian / Keterangan Transaksi": "Takziah ke rumah B6 No. 03 ( anaknya meninggal dunia)", "Debet (Masuk)": 0.0, "Kredit (Keluar)": 100000.0}
    ])
    st.session_state.df_kas_sosial_state = muat_data_kas(FILE_KAS_SOSIAL, default_sosial)

def parsing_angka_aman(val):
    if pd.isna(val) or val == "" or val == "-":
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    val_str = str(val).strip().replace("Rp", "").replace(" ", "")
    clean_str = "".join([c for c in val_str if c.isdigit() or c == '-' or c == '.' or c == ','])
    if not clean_str or clean_str == "-":
        return 0.0
    if ',' in clean_str and '.' in clean_str:
        clean_str = clean_str.replace('.', '').replace(',', '.')
    elif ',' in clean_str:
        clean_str = clean_str.replace(',', '')
    elif clean_str.count('.') > 1:
        parts = clean_str.split('.')
        clean_str = "".join(parts[:-1]) + "." + parts[-1]
    try:
        return float(clean_str)
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
    if df_input.empty:
        return pd.DataFrame(columns=["Tanggal", "Uraian / Keterangan Transaksi", "Debet (Masuk)", "Kredit (Keluar)", "Saldo (Rp)"])
        
    df = df_input.copy()
    df['_dt_sort'] = pd.to_datetime(df['Tanggal'], format='%d/%m/%Y', errors='coerce')
    df = df.sort_values(by='_dt_sort').drop(columns=['_dt_sort']).reset_index(drop=True)

    curr = 0.0
    tanggal_list, uraian_list, debet_val, kredit_val, saldo_formatted = [], [], [], [], []
    
    for idx, row in df.iterrows():
        tgl = str(row.get("Tanggal", ""))
        uraian = str(row.get("Uraian / Keterangan Transaksi", ""))
        deb = parsing_angka_aman(row.get("Debet (Masuk)", 0))
        kre = parsing_angka_aman(row.get("Kredit (Keluar)", 0))
        
        if idx == 0:
            curr = deb - kre
        else:
            curr = curr + deb - kre
            
        tanggal_list.append(tgl)
        uraian_list.append(uraian)
        debet_val.append(deb)
        kredit_val.append(kre)
        saldo_formatted.append(format_Rupiah(curr))
        
    return pd.DataFrame({
        "Tanggal": tanggal_list,
        "Uraian / Keterangan Transaksi": uraian_list,
        "Debet (Masuk)": debet_val,
        "Kredit (Keluar)": kredit_val,
        "Saldo (Rp)": saldo_formatted
    })

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
            new_data = [(255, 255, 255, 0) if item[0] > 240 and item[1] > 240 and item[2] > 240 else item for item in datas]
            img.putdata(new_data)
            st.image(img, width=380)
        except:
            st.image(logo_path, width=380)
    else:
        st.write("🏛️")

with col_title:
    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown('<div class="jumbo-title">SISTEM INFORMASI ADMINISTRASI KEPENDUDUKAN (SIAK)</div>', unsafe_allow_html=True)
    st.markdown('<div class="jumbo-subtitle">Rukun Tetangga 06 / Rukun Warga 14 • Perum Griya Permata Raya • Desa Nanjung Mekar, Rancaekek</div>', unsafe_allow_html=True)

st.write("---")

def load_data_rt06_stable():
    if not os.path.exists(FILE_EXCEL_WARGA):
        return pd.DataFrame()
    
    try:
        df = pd.read_excel(FILE_EXCEL_WARGA, header=3, dtype=str)
        df.columns = df.columns.astype(str).str.strip().str.upper()
        
        df = df.rename(columns={"STUS RUMAH": "STATUS RUMAH"})
        
        cols = list(df.columns)
        rumah_idx = -1
        for i, c in enumerate(cols):
            if "RUMAH" in c or "ALAMAT" in c:
                rumah_idx = i
                break
        if rumah_idx != -1 and len(cols) > rumah_idx + 1:
            cols[rumah_idx + 1] = "NAMA KEPALA KELUARGA"
            df.columns = cols
        
        df = df.loc[:, ~df.columns.str.contains('UNNAMED')]
        df = df.dropna(how="all")
        
        for col_name in list(df.columns):
            if col_name == "NO" or col_name == "NO.":
                df = df.drop(columns=[col_name])

        for col in df.select_dtypes(include=['object']).columns:
            df[col] = df[col].astype(str).str.strip()
            df.loc[df[col].str.lower() == 'nan', col] = None
        
        bulan_indo = {
            1: "Januari", 2: "Februari", 3: "Maret", 4: "April", 5: "Mei", 6: "Juni",
            7: "Juli", 8: "Agustus", 9: "September", 10: "Oktober", 11: "November", 12: "Desember"
        }
        
        for c in df.columns:
            if "TGL" in c or "TANGGAL" in c or "LAHIR" in c:
                def format_tgl_bersih(val):
                    if pd.isnull(val) or str(val).lower() in ['nan', 'none', '']:
                        return ""
                    val_str = str(val).strip()
                    if "00:00:00" in val_str:
                        val_str = val_str.replace("00:00:00", "").strip()
                    try:
                        dt = pd.to_datetime(val_str, errors='coerce')
                        if pd.notnull(dt):
                            return f"{dt.day:02d} {bulan_indo.get(dt.month, '')} {dt.year}"
                    except:
                        pass
                    return val_str
                df[c] = df[c].apply(format_tgl_bersih)

        df['_ORIGINAL_IDX'] = df.index

        col_rumah_sort = next((col for col in df.columns if "RUMAH" in col or "ALAMAT" in col), None)
        col_kk_sort = next((col for col in df.columns if "KEPALA" in col or "KK" in col), None)
        
        if col_rumah_sort:
            df['_TEMP_RUMAH'] = df[col_rumah_sort].ffill()
            if col_kk_sort:
                df['_TEMP_KK'] = df[col_kk_sort].ffill()
                df = df.sort_values(by=['_TEMP_RUMAH', '_TEMP_KK'], key=lambda x: x.astype(str).str.lower()).reset_index(drop=True)
            else:
                df = df.sort_values(by='_TEMP_RUMAH', key=lambda x: x.astype(str).str.lower()).reset_index(drop=True)
            df = df.drop(columns=[c for c in df.columns if c.startswith('_TEMP_')])
            
        return df
    except Exception as e:
        st.error(f"Gagal memuat database kependudukan: {e}")
        return pd.DataFrame()

st.session_state.df_warga_state = load_data_rt06_stable()
df = st.session_state.df_warga_state

if not df.empty:
    col_kk_candi = [c for c in df.columns if "KEPALA" in c or "KK" in c]
    col_kk = col_kk_candi[0] if col_kk_candi else df.columns[1]
    
    col_nama_candi = [c for c in df.columns if "ANGGOTA" in c or "NAMA" in c]
    col_nama = col_nama_candi[0] if col_nama_candi else (df.columns[2] if len(df.columns) > 2 else df.columns[0])
    
    col_rumah_candi = [c for c in df.columns if "RUMAH" in c or "ALAMAT" in c]
    col_rumah = col_rumah_candi[0] if col_rumah_candi else None

    col_jk = next((c for c in df.columns if "JK" in c or "KELAMIN" in c or "GENDER" in c), None)
    col_usia = next((c for c in df.columns if "USIA" in c or "UMUR" in c), None)

    daftar_blok_lengkap = [
        "B3-01", "B3-02", "B3-03", "B3-04", "B3-05", "B3-06", "B3-07", "B3-08", "B3-09", "B3-10",
        "B3-11", "B3-12", "B3-13", "B3-14", "B3-15", "B3-16", "B3-17", "B3-18", "B3-19", "B3-20",
        "B4-01", "B4-02", "B4-03", "B4-04", "B4-05", "B4-06", "B4-07", "B4-08", "B4-09", "B4-10"
    ]
    if col_rumah and not df.empty:
        df_temp_b = df.copy()
        df_temp_b['_TEMP_R'] = df_temp_b[col_rumah].ffill()
        existing_blocks = sorted(list(set(df_temp_b['_TEMP_R'].dropna().astype(str).tolist())))
        for eb in existing_blocks:
            if eb not in daftar_blok_lengkap and eb.lower() != 'nan' and eb.strip() != '':
                daftar_blok_lengkap.append(eb)

    st.sidebar.markdown("### 🧭 Navigasi Menu Administrasi")
    daftar_menu_pilihan = [
        "Dashboard Eksekutif Kependudukan",
        "📋 Database Kependudukan & Demografi",
        "✏️ Pemutakhiran & Koreksi Data Penduduk",
        "🗂️ Layanan Arsip Kartu Keluarga (KK)", 
        "📈 Analisis & Statistik Demografi", 
        "📊 Laporan Rekapitulasi Administrasi",
        "💰 Administrasi Keuangan RT & Sosial",
        "🖨️ Pusat Dokumen & Ekspor Laporan"
    ]

    def update_menu_pilihan():
        st.session_state.selected_menu = st.session_state.widget_nav_selectbox

    selected_sidebar = st.sidebar.selectbox(
        "Pilih Modul Layanan:", 
        daftar_menu_pilihan, 
        index=daftar_menu_pilihan.index(st.session_state.selected_menu) if st.session_state.selected_menu in daftar_menu_pilihan else 0,
        key="widget_nav_selectbox",
        on_change=update_menu_pilihan
    )

    menu = st.session_state.selected_menu

    if menu == "Dashboard Eksekutif Kependudukan":
        col_jam1, col_jam2 = st.columns([2, 2])
        with col_jam1:
            st.subheader("📊 Dashboard Eksekutif Kependudukan")
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
            jk_series = df[col_jk].fillna("").astype(str).str.upper().str.strip()
            jml_p = len(df[jk_series.str.contains("P", regex=False)])
            jml_l = total_jiwa - jml_p
        else:
            jml_l = total_jiwa
            jml_p = 0

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
            <div class="metric-card" style="border-left: 6px solid #0d9488;">
                <div class="metric-title">Total Jiwa</div>
                <div class="metric-value" style="color: #0f766e;">{total_jiwa} <span style="font-size: 16px; font-weight: 600; color: #64748b;">Jiwa</span></div>
            </div>
            <div class="metric-card" style="border-left: 6px solid #059669;">
                <div class="metric-title">Laki-laki</div>
                <div class="metric-value" style="color: #065f46;">{jml_l} <span style="font-size: 16px; font-weight: 600; color: #64748b;">Orang</span></div>
            </div>
            <div class="metric-card" style="border-left: 6px solid #0284c7;">
                <div class="metric-title">Perempuan</div>
                <div class="metric-value" style="color: #0369a1;">{jml_p} <span style="font-size: 16px; font-weight: 600; color: #64748b;">Orang</span></div>
            </div>
            <div class="metric-card" style="border-left: 6px solid #db2777;">
                <div class="metric-title">Balita (0-5 th)</div>
                <div class="metric-value" style="color: #9d174d;">{jml_balita} <span style="font-size: 16px; font-weight: 600; color: #64748b;">Jiwa</span></div>
            </div>
            <div class="metric-card" style="border-left: 6px solid #d97706;">
                <div class="metric-title">Lansia (>60 th)</div>
                <div class="metric-value" style="color: #b45309;">{jml_lansia} <span style="font-size: 16px; font-weight: 600; color: #64748b;">Jiwa</span></div>
            </div>
        </div>
        """, unsafe_allow_html=True)

        st.write("---")
        st.markdown("### 🚀 Panel Navigasi Utama SIAK")
        
        col_m1, col_m2 = st.columns(2)
        with col_m1:
            if st.button("📋 Database Kependudukan & Demografi", use_container_width=True, key="btn_m1"):
                st.session_state.selected_menu = "📋 Database Kependudukan & Demografi"
                st.rerun()
            if st.button("✏️ Pemutakhiran & Koreksi Data Penduduk", use_container_width=True, key="btn_m1_edit"):
                st.session_state.selected_menu = "✏️ Pemutakhiran & Koreksi Data Penduduk"
                st.rerun()
            if st.button("🗂️ Layanan Arsip Kartu Keluarga (KK)", use_container_width=True, key="btn_m2"):
                st.session_state.selected_menu = "🗂️ Layanan Arsip Kartu Keluarga (KK)"
                st.rerun()
            if st.button("📊 Laporan Rekapitulasi Administrasi", use_container_width=True, key="btn_m3"):
                st.session_state.selected_menu = "📊 Laporan Rekapitulasi Administrasi"
                st.rerun()
        with col_m2:
            if st.button("📈 Analisis & Statistik Demografi", use_container_width=True,
