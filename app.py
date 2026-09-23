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
    page_title="Portal Resmi RT 06 RW 14 Griya Permata Raya",
    layout="wide",
    page_icon="🏠"
)

# Custom CSS Standar Profesional, Terang, & Jelas Dibaca
st.markdown("""
<style>
    .stApp {
        background: linear-gradient(135deg, #f8fafc 0%, #cbd5e1 100%) !important;
    }
    .main .block-container {
        background: #ffffff !important;
        padding: 2.5rem 3rem !important;
        border-radius: 20px !important;
        box-shadow: 0 10px 30px rgba(0, 0, 0, 0.08) !important;
        margin-top: 1.5rem;
        margin-bottom: 1.5rem;
    }
    h1, h2, h3, h4, h5, h6 {
        color: #0f172a !important;
        font-weight: 800 !important;
    }
    p, label, span, div {
        color: #334155 !important;
    }
    .metric-card {
        background: #f1f5f9;
        border: 1px solid #cbd5e1;
        padding: 20px;
        border-radius: 14px;
        box-shadow: 0 4px 12px rgba(0, 0, 0, 0.03);
        margin-bottom: 12px;
    }
    .metric-title {
        font-size: 12px !important;
        text-transform: uppercase;
        font-weight: 700;
        color: #475569;
        letter-spacing: 0.5px;
    }
    .metric-value {
        font-size: 28px !important;
        font-weight: 900 !important;
        margin-top: 6px;
        color: #1e3a8a;
    }
    .jumbo-title {
        font-size: 54px !important;
        font-weight: 900 !important;
        color: #1e3a8a !important;
        line-height: 1.1 !important;
        margin-bottom: 6px !important;
    }
    .jumbo-subtitle {
        font-size: 22px !important;
        font-weight: 700 !important;
        color: #475569 !important;
    }
    .stButton button {
        font-size: 16px !important;
        font-weight: 700 !important;
        padding: 16px 20px !important;
        border-radius: 12px !important;
        border: none !important;
        color: #ffffff !important;
        -webkit-text-fill-color: #ffffff !important;
        box-shadow: 0 4px 12px rgba(37, 99, 235, 0.2) !important;
        transition: all 0.3s ease !important;
        width: 100% !important;
        background: linear-gradient(135deg, #2563eb, #1d4ed8) !important;
    }
    .stButton button:hover {
        transform: translateY(-2px) !important;
        box-shadow: 0 6px 16px rgba(37, 99, 235, 0.3) !important;
    }
    section[data-testid="stSidebar"] {
        background-color: #1e293b !important;
    }
    section[data-testid="stSidebar"] * {
        color: #f8fafc !important;
    }
</style>
""", unsafe_allow_html=True)

if 'selected_menu' not in st.session_state:
    st.session_state.selected_menu = "Beranda / Dashboard"

FILE_KAS_RT = "penyimpanan_kas_rt.csv"
FILE_KAS_SOSIAL = "penyimpanan_kas_sosial.csv"
FILE_EXCEL_WARGA = "data_warga_rt06.xlsx"

def muat_data_kas(file_path, default_df):
    if os.path.exists(file_path):
        try:
            df_disk = pd.read_csv(file_path)
            if not df_disk.empty:
                return df_disk
        except:
            pass
    return default_df

def simpan_data_kas(file_path, df):
    try:
        df.to_csv(file_path, index=False)
    except:
        pass

if 'df_kas_rt_state' not in st.session_state:
    default_rt = pd.DataFrame(columns=["Tanggal", "Uraian / Keterangan Transaksi", "Debet (Masuk)", "Kredit (Keluar)"])
    st.session_state.df_kas_rt_state = muat_data_kas(FILE_KAS_RT, default_rt)

if 'df_kas_sosial_state' not in st.session_state:
    default_sosial = pd.DataFrame(columns=["Tanggal", "Uraian / Keterangan Transaksi", "Debet (Masuk)", "Kredit (Keluar)"])
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
col_logo, col_title = st.columns([1.2, 4])
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
            st.image(img, width=420)
        except:
            st.markdown("<h1>🏠</h1>", unsafe_allow_html=True)
    else:
        st.markdown("<h1>🏠</h1>", unsafe_allow_html=True)

with col_title:
    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown('<div class="jumbo-title">PORTAL RESMI RT 06 / RW 14</div>', unsafe_allow_html=True)
    st.markdown('<div class="jumbo-subtitle">Griya Permata Raya • Desa Nanjung Mekar, Rancaekek</div>', unsafe_allow_html=True)

st.write("---")

def load_data_rt06_stable():
    if not os.path.exists(FILE_EXCEL_WARGA):
        return pd.DataFrame()
    
    try:
        df = pd.read_excel(FILE_EXCEL_WARGA, header=3)
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

        col_usia_candi = next((c for c in df.columns if "USIA" in c or "UMUR" in c), None)
        col_tgl_lahir_candi = next((c for c in df.columns if "LAHIR" in c and ("TGL" in c or "TANGGAL" in c)), None)
        
        if col_usia_candi:
            current_date = datetime.now(ZoneInfo("Asia/Jakarta"))
            for idx_u, row_u in df.iterrows():
                usia_final = 30
                if col_tgl_lahir_candi:
                    tgl_str = str(row_u.get(col_tgl_lahir_candi, ""))
                    try:
                        dt_lahir = pd.to_datetime(tgl_str, errors='coerce')
                        if pd.notnull(dt_lahir):
                            calc_age = current_date.year - dt_lahir.year - ((current_date.month, current_date.day) < (dt_lahir.month, dt_lahir.day))
                            if 0 <= calc_age <= 120:
                                usia_final = int(calc_age)
                            else:
                                raise ValueError()
                        else:
                            raise ValueError()
                    except:
                        found_y = False
                        for p_str in tgl_str.split():
                            if p_str.isdigit() and len(p_str) == 4 and 1900 <= int(p_str) <= 2026:
                                calc_age_y = 2026 - int(p_str)
                                if 0 <= calc_age_y <= 120:
                                    usia_final = int(calc_age_y)
                                    found_y = True
                                    break
                        if not found_y:
                            try:
                                val_raw = row_u.get(col_usia_candi)
                                num_u = int(float(str(val_raw).strip()))
                                if 0 <= num_u <= 120:
                                    usia_final = int(num_u)
                            except:
                                usia_final = 30
                else:
                    usia_final = 30
                
                # DIKUNCI JADI INTEGER LALU KE STRING AGAR TIDAK MUNCUL .000000
                df.loc[idx_u, col_usia_candi] = str(int(usia_final))

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
        st.error(f"Gagal memuat data warga: {e}")
        return pd.DataFrame()

if 'df_warga_state' not in st.session_state:
    st.session_state.df_warga_state = load_data_rt06_stable()

df = st.session_state.df_warga_state

if df.empty:
    st.warning("⚠️ File Excel data warga belum ditemukan atau kosong. Pastikan file 'data_warga_rt06.xlsx' tersedia.")
else:
    col_kk_candi = [c for c in df.columns if "KEPALA" in c or "KK" in c]
    col_kk = col_kk_candi[0] if col_kk_candi else df.columns[1]
    
    col_nama_candi = [c for c in df.columns if "ANGGOTA" in c or "NAMA" in c]
    col_nama = col_nama_candi[0] if col_nama_candi else (df.columns[2] if len(df.columns) > 2 else df.columns[0])
    
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
        "📊 Rekapitulasi Administrasi RW",
        "💰 Laporan Kas RT & Sosial (Perelek R6 Sauyunan)",
        "🖨️ Cetak Rekap PDF"
    ]

    def update_menu():
        st.session_state.selected_menu = st.session_state.sb_menu

    selected_sidebar = st.sidebar.selectbox(
        "Pilih Halaman:", 
        daftar_menu_pilihan, 
        index=daftar_menu_pilihan.index(st.session_state.selected_menu) if st.session_state.selected_menu in daftar_menu_pilihan else 0,
        key="sb_menu",
        on_change=update_menu
    )
    menu = st.session_state.selected_menu

    if menu == "Beranda / Dashboard":
        col_jam1, col_jam2 = st.columns([2, 2])
        with col_jam1:
            st.subheader("📊 Dashboard Eksekutif Warga")
        with col_jam2:
            placeholder_waktu = st.empty()

        df_ffill = df.copy()
        if col_kk:
            df_ffill[col_kk] = df_ffill[col_kk].replace('', pd.NA).ffill()
        if col_rumah:
            df_ffill[col_rumah] = df_ffill[col_rumah].replace('', pd.NA).ffill()

        total_jiwa = len(df)
        total_kk = df_ffill[col_kk].nunique() if col_kk in df_ffill.columns else 0
        
        jml_l = len(df[df[col_jk].astype(str).str.upper().str.contains("L")]) if col_jk else 0
        jml_p = len(df[df[col_jk].astype(str).str.upper().str.contains("P")]) if col_jk else 0

        jml_balita, jml_lansia = 0, 0
        if col_usia:
            try:
                u_ser = df[col_usia].apply(lambda x: int(float(str(x))))
                jml_balita = len(u_ser[(u_ser >= 0) & (u_ser <= 5)])
                jml_lansia = len(u_ser[u_ser > 60])
            except:
                pass

        st.markdown(f"""
        <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(160px, 1fr)); gap: 16px; margin-top: 10px; margin-bottom: 25px;">
            <div class="metric-card" style="border-left: 5px solid #2563eb;">
                <div class="metric-title">Jumlah KK</div>
                <div class="metric-value">{total_kk} <span style="font-size: 14px; color: #64748b;">KK</span></div>
            </div>
            <div class="metric-card" style="border-left: 5px solid #059669;">
                <div class="metric-title">Total Jiwa</div>
                <div class="metric-value" style="color: #065f46;">{total_jiwa} <span style="font-size: 14px; color: #64748b;">Jiwa</span></div>
            </div>
            <div class="metric-card" style="border-left: 5px solid #0284c7;">
                <div class="metric-title">Laki-laki</div>
                <div class="metric-value" style="color: #0369a1;">{jml_l}</div>
            </div>
            <div class="metric-card" style="border-left: 5px solid #db2777;">
                <div class="metric-title">Perempuan</div>
                <div class="metric-value" style="color: #9d174d;">{jml_p}</div>
            </div>
            <div class="metric-card" style="border-left: 5px solid #d97706;">
                <div class="metric-title">Balita (0-5 th)</div>
                <div class="metric-value" style="color: #b45309;">{jml_balita}</div>
            </div>
            <div class="metric-card" style="border-left: 5px solid #7c3aed;">
                <div class="metric-title">Lansia (>60 th)</div>
                <div class="metric-value" style="color: #5b21b6;">{jml_lansia}</div>
            </div>
        </div>
        """, unsafe_allow_html=True)

        st.write("---")
        st.markdown("### 🚀 Menu Utama Portal RT 06")
        
        col_m1, col_m2 = st.columns(2)
        with col_m1:
            if st.button("📋 Data Seluruh Warga", use_container_width=True, key="btn_main_warga_v10"):
                st.session_state.selected_menu = "📋 Data Seluruh Warga"
                st.rerun()
            if st.button("🗂️ Cetak Kartu Keluarga (KK)", use_container_width=True, key="btn_main_kk_v10"):
                st.session_state.selected_menu = "🗂️ Cetak Kartu Keluarga (KK)"
                st.rerun()
            if st.button("📊 Rekapitulasi Administrasi RW", use_container_width=True, key="btn_main_rw_v10"):
                st.session_state.selected_menu = "📊 Rekapitulasi Administrasi RW"
                st.rerun()
        with col_m2:
            if st.button("📈 Grafik Demografi", use_container_width=True, key="btn_main_grafik_v10"):
                st.session_state.selected_menu = "📈 Grafik Demografi"
                st.rerun()
            if st.button("💰 Laporan Kas RT & Sosial", use_container_width=True, key="btn_main_kas_v10"):
                st.session_state.selected_menu = "💰 Laporan Kas RT & Sosial (Perelek R6 Sauyunan)"
                st.rerun()
            if st.button("🖨️ Cetak Laporan Rekap PDF", use_container_width=True, key="btn_main_pdf_v10"):
                st.session_state.selected_menu = "🖨️ Cetak Rekap PDF"
                st.rerun()

        for _ in range(5):
            waktu_sekarang = datetime.now(ZoneInfo("Asia/Jakarta"))
            bulan_indo_nama = {1: "Januari", 2: "Februari", 3: "Maret", 4: "April", 5: "Mei", 6: "Juni", 7: "Juli", 8: "Agustus", 9: "September", 10: "Oktober", 11: "November", 12: "Desember"}
            tgl_str = f"{waktu_sekarang.day:02d} {bulan_indo_nama.get(waktu_sekarang.month, '')} {waktu_sekarang.year}"
            jam_str = waktu_sekarang.strftime("%H:%M:%S")
            
            placeholder_waktu.markdown(f"""
            <div style="background: #f8fafc; border: 1px solid #cbd5e1; padding: 10px 16px; border-radius: 12px; text-align: right;">
                <span style="font-size: 11px; color: #64748b; font-weight: 700;">🕒 Live Update (WIB):</span><br>
                <strong style="font-size: 13px; color: #0f172a;">{tgl_str} | {jam_str} WIB</strong>
            </div>
            """, unsafe_allow_html=True)
            time.sleep(1)
        st.rerun()

    elif menu == "📋 Data Seluruh Warga":
        if st.button("⬅️ Kembali ke Beranda", key="back_from_warga_v10"):
            st.session_state.selected_menu = "Beranda / Dashboard"
            st.rerun()
        st.subheader("📋 Data Keseluruhan Warga, Hapus Warga, & Input Stabil")
        
        def highlight_luar_nm(row):
            row_str = str(row.values).lower()
            if "luar nm" in row_str:
                return ['background-color: #fef08a'] * len(row)
            return [''] * len(row)

        try:
            st.dataframe(df.style.apply(highlight_luar_nm, axis=1), use_container_width=True, hide_index=True)
        except:
            st.dataframe(df, use_container_width=True, hide_index=True)

        st.markdown("---")
        st.markdown("### 🗑️ Hapus Data Warga (Pindah / Keluar)")
        
        with st.form("form_hapus_warga_stabil_bawah_v10", clear_on_submit=False):
            list_warga_pilih = [f"Baris {i+1} | Rumah: {row.get(col_rumah, '-')} | KK: {row.get(col_kk, '-')} | Nama: {row.get(col_nama, '-')}" for i, row in df.iterrows()]
            target_hapus_str = st.selectbox("Pilih Warga yang Ingin Dihapus:", ["(Pilih warga...)"] + list_warga_pilih, key="sel_hapus_bawah_v10")
            
            btn_eksekusi_hapus = st.form_submit_button("🗑️ Hapus Data Warga Ini Secara Permanen")
            
            if btn_eksekusi_hapus:
                if target_hapus_str == "(Pilih warga...)":
                    st.warning("⚠️ Silakan pilih data warga terlebih dahulu dari dropdown.")
                else:
                    try:
                        idx_hapus = int(target_hapus_str.split("|")[0].replace("Baris", "").strip()) - 1
                        df_setelah_hapus = df.drop(index=idx_hapus).reset_index(drop=True)

                        import openpyxl
                        wb = openpyxl.Workbook()
                        ws = wb.active
                        ws.title = "Data Warga"
                        ws.append(["DATA WARGA RT 06 RW 14"])
                        ws.append([])
                        ws.append([])
                        ws.append(list(df_setelah_hapus.columns))
                        for _, r in df_setelah_hapus.iterrows():
                            ws.append(list(r.values))
                        wb.save(FILE_EXCEL_WARGA)

                        st.session_state.df_warga_state = load_data_rt06_stable()
                        st.success("✅ Data warga berhasil dihapus secara permanen!")
                        time.sleep(1)
                        st.rerun()
                    except Exception as e:
                        st.error(f"❌ Gagal menghapus data: {e}")

        st.markdown("---")
        st.markdown("### ➕ Form Input Warga Baru / Anggota Keluarga (Lengkap dengan Kolom Usia Otomatis)")

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

        with st.form("form_input_warga_lengkap_v6", clear_on_submit=False):
            col_f1, col_f2 = st.columns(2)
            with col_f1:
                in_no_rumah = st.selectbox("No. Rumah", sorted(list(set(daftar_blok_lengkap))), key="in_no_rumah_full_v6")
                in_nama_kk = st.text_input("NAMA KEPALA KELUARGA", key="in_nama_kk_full_v6")
                in_nama_anggota = st.text_input("Nama Lengkap Anggota Keluarga", key="in_nama_anggota_full_v6")
                in_jk = st.selectbox("Jenis Kelamin", ["L", "P"], key="in_jk_full_v6")
                in_hub = st.selectbox("Hubungan Keluarga", ["Kepala Keluarga", "Istri", "Anak Kandung", "Famili Lain", "Mertua"], key="in_hub_full_v6")
                in_tmplhr = st.text_input("Tempat Lahir", key="in_tmplhr_full_v6")
                
                st.markdown("📅 **Tanggal Lahir (Tgl, Bulan, Tahun):**")
                c_t1, c_t2, c_t3 = st.columns(3)
                with c_t1: in_tgl = st.number_input("Tgl", min_value=1, max_value=31, value=1, key="in_tgl_v6")
                with c_t2: in_bln = st.number_input("Bln", min_value=1, max_value=12, value=1, key="in_bln_v6")
                with c_t3: in_thn = st.number_input("Thn", min_value=1900, max_value=2026, value=1995, key="in_thn_v6")
                
                bulan_indo_nama = {1: "Januari", 2: "Februari", 3: "Maret", 4: "April", 5: "Mei", 6: "Juni", 7: "Juli", 8: "Agustus", 9: "September", 10: "Oktober", 11: "November", 12: "Desember"}
                tanggal_lahir_lengkap_str = f"{int(in_tgl):02d} {bulan_indo_nama.get(int(in_bln), 'Januari')} {int(in_thn)}"
                
                current_dt = datetime.now(ZoneInfo("Asia/Jakarta"))
                usia_otomatis = int(current_dt.year - int(in_thn) - ((current_dt.month, current_dt.day) < (int(in_bln), int(in_tgl))))
                if usia_otomatis < 0: usia_otomatis = 0
                
                st.info(f"💡 Tgl Lahir: **{tanggal_lahir_lengkap_str}** | Usia Otomatis Terisi: **{usia_otomatis} Tahun**")

            with col_f2:
                in_status_nikah = st.selectbox("Status Perkawinan", ["Belum Kawin", "Kawin", "Cerai Hidup", "Cerai Mati"], key="in_status_nikah_full_v6")
                in_agama = st.selectbox("Agama", ["Islam", "Kristen", "Katolik", "Hindu", "Buddha", "Konghucu"], key="in_agama_full_v6")
                in_goldarah = st.selectbox("Gol. Darah", ["A", "B", "AB", "O", "Tidak Tahu"], key="in_goldarah_full_v6")
                in_suku = st.selectbox("Etnis / Suku", ["Sunda", "Jawa", "Padang", "Batak", "Betawi", "Lainnya"], key="in_suku_full_v6")
                in_pend = st.selectbox("Pendidikan", ["Tamat SLTA/sederajat", "Tamat SLTP/sederajat", "Tamat SD/sederajat", "Diploma IV / Strata I", "Sedang SD/sedajerat", "Sedang SLTP/sederajat", "Belum / Tidak Sekolah"], key="in_pend_full_v6")
                in_pek = st.selectbox("Pekerjaan", ["Karyawan Swasta", "Wiraswasta", "Mengurus Rumah Tangga", "Belum Bekerja", "Pelajar", "PNS / TNI / Polri"], key="in_pek_full_v6")
                in_status_rumah = st.selectbox("Status Rumah", ["Milik / Tetap", "Sewa/Kontrak", "Kosong"], key="in_status_rumah_full_v6")
                in_status_domisili = st.selectbox("Status Domisili", ["Nanjung Mekar", "luar NM"], key="in_status_domisili_full_v6")

            if st.form_submit_button("💾 Masukkan Data ke Database Excel"):
                try:
                    df_raw_excel = pd.read_excel(FILE_EXCEL_WARGA, header=3)
                    df_raw_excel.columns = df_raw_excel.columns.astype(str).str.strip().str.upper()
                    df_raw_excel = df_raw_excel.rename(columns={"STUS RUMAH": "STATUS RUMAH"})
                    cols_ex = list(df_raw_excel.columns)
                    r_idx = -1
                    for i, c in enumerate(cols_ex):
                        if "RUMAH" in c or "ALAMAT" in c:
                            r_idx = i
                            break
                    if r_idx != -1 and len(cols_ex) > r_idx + 1:
                        cols_ex[r_idx + 1] = "NAMA KEPALA KELUARGA"
                        df_raw_excel.columns = cols_ex

                    baris_baru_dict = {}
                    for col_excel in df_raw_excel.columns:
                        c_up = str(col_excel).upper()
                        if col_excel == "NO" or col_excel == "NO.":
                            continue
                        elif ("RUMAH" in c_up and "STATUS" not in c_up) or "ALAMAT" in c_up:
                            baris_baru_dict[col_excel] = in_no_rumah if in_hub.lower() == "kepala keluarga" else None
                        elif "KEPALA" in c_up or "KK" in c_up:
                            baris_baru_dict[col_excel] = in_nama_kk
                        elif "ANGGOTA" in c_up or "NAMA" in c_up:
                            baris_baru_dict[col_excel] = in_nama_anggota
                        elif "JK" in c_up or "KELAMIN" in c_up:
                            baris_baru_dict[col_excel] = in_jk
                        elif "HUBUNGAN" in c_up:
                            baris_baru_dict[col_excel] = in_hub
                        elif "TEMPAT" in c_up:
                            baris_baru_dict[col_excel] = in_tmplhr
                        elif "LAHIR" in c_up and ("TGL" in c_up or "TANGGAL" in c_up):
                            baris_baru_dict[col_excel] = tanggal_lahir_lengkap_str
                        elif "USIA" in c_up or "UMUR" in c_up:
                            baris_baru_dict[col_excel] = int(usia_otomatis)
                        elif c_up == "STATUS" or "KAWIN" in c_up or "NIKAH" in c_up:
                            baris_baru_dict[col_excel] = in_status_nikah
                        elif "AGAMA" in c_up:
                            baris_baru_dict[col_excel] = in_agama
                        elif "DARAH" in c_up or "GOL" in c_up:
                            baris_baru_dict[col_excel] = in_goldarah
                        elif "SUKU" in c_up or "ETNIS" in c_up:
                            baris_baru_dict[col_excel] = in_suku
                        elif "PENDIDIKAN" in c_up:
                            baris_baru_dict[col_excel] = in_pend
                        elif "PEKERJAAN" in c_up:
                            baris_baru_dict[col_excel] = in_pek
                        elif "STATUS RUMAH" in c_up:
                            baris_baru_dict[col_excel] = in_status_rumah
                        elif "STATUS DOMISILI" in c_up:
                            baris_baru_dict[col_excel] = in_status_domisili
                        else:
                            baris_baru_dict[col_excel] = "-"

                    col_kk_target = next((col for col in df_raw_excel.columns if "KEPALA" in col or "KK" in col), None)
                    insert_idx = len(df_raw_excel)
                    if col_kk_target:
                        for idx_r, row_val in df_raw_excel.iterrows():
                            if str(row_val.get(col_kk_target, "")).strip().lower() == in_nama_kk.strip().lower():
                                insert_idx = idx_r + 1

                    df_updated_excel = pd.concat([df_raw_excel.iloc[:insert_idx], pd.DataFrame([baris_baru_dict]), df_raw_excel.iloc[insert_idx:]], ignore_index=True)
                    
                    import openpyxl
                    wb = openpyxl.Workbook()
                    ws = wb.active
                    ws.title = "Data Warga"
                    ws.append(["DATA WARGA RT 06 RW 14"])
                    ws.append([])
                    ws.append([])
                    ws.append(list(df_updated_excel.columns))
                    for _, r in df_updated_excel.iterrows():
                        ws.append(list(r.values))
                    wb.save(FILE_EXCEL_WARGA)

                    st.session_state.df_warga_state = load_data_rt06_stable()
                    st.success("✅ Data warga baru berhasil disimpan!")
                    time.sleep(1)
                    st.rerun()
                except Exception as e:
                    st.error(f"❌ Gagal menyimpan data: {e}")

    elif menu == "🗂️ Cetak Kartu Keluarga (KK)":
        if st.button("⬅️ Kembali ke Beranda", key="back_from_kk_v10"):
            st.session_state.selected_menu = "Beranda / Dashboard"
            st.rerun()
        st.subheader("🗂️ Cetak Kartu Keluarga (KK)")
        
        df_ffill = df.copy()
        if col_kk:
            df_ffill[col_kk] = df_ffill[col_kk].replace('', pd.NA).ffill()
        if col_rumah:
            df_ffill[col_rumah] = df_ffill[col_rumah].replace('', pd.NA).ffill()

        daftar_kk = sorted(list(set([str(x).strip() for x in df_ffill[col_kk].dropna().tolist() if str(x).strip().lower() not in ['', 'nan', 'none']])))
        pilihan_kk = st.selectbox("Pilih Kepala Keluarga:", daftar_kk, key="select_kk_main_v10")
        
        if pilihan_kk:
            df_keluarga = df_ffill[df_ffill[col_kk].astype(str).str.strip().str.lower() == pilihan_kk.strip().lower()].copy()
            no_rmh = str(df_keluarga[col_rumah].dropna().iloc[0]) if col_rumah and not df_keluarga[df_keluarga[col_rumah].notna()].empty else "-"
            cols_tampilan_web = [c for c in df_keluarga.columns if c != col_kk and "URUT" not in c and c != "NO"]
            
            st.markdown(f"**No. Rumah:** {no_rmh} | **Kepala Keluarga:** {pilihan_kk}")
            st.dataframe(df_keluarga[cols_tampilan_web], use_container_width=True, hide_index=True)

    elif menu == "📈 Grafik Demografi":
        if st.button("⬅️ Kembali ke Beranda", key="back_from_grafik_v10"):
            st.session_state.selected_menu = "Beranda / Dashboard"
            st.rerun()
        st.subheader("📈 Analisis Grafik Demografi Warga")
        if col_jk:
            df_jk = df[col_jk].dropna().value_counts().reset_index()
            df_jk.columns = ["Jenis Kelamin", "Jumlah"]
            fig_jk = px.pie(df_jk, names="Jenis Kelamin", values="Jumlah", hole=0.5, title="👥 Rasio Berdasarkan Jenis Kelamin")
            st.plotly_chart(fig_jk, use_container_width=True, key="chart_jk_main_v10")

    elif menu == "📊 Rekapitulasi Administrasi RW":
        if st.button("⬅️ Kembali ke Beranda", key="back_from_rw_v10"):
            st.session_state.selected_menu = "Beranda / Dashboard"
            st.rerun()
        st.subheader("📊 Rekapitulasi Administrasi RW")
        total_kk_rw = df[col_kk].nunique() if col_kk in df.columns else 0
        total_jiwa_rw = len(df)
        st.write(f"- Jumlah KK: {total_kk_rw}")
        st.write(f"- Total Jiwa: {total_jiwa_rw}")

    elif menu == "💰 Laporan Kas RT & Sosial (Perelek R6 Sauyunan)":
        if st.button("⬅️ Kembali ke Beranda", key="back_from_kas_v10"):
            st.session_state.selected_menu = "Beranda / Dashboard"
            st.rerun()
        st.subheader("💰 Laporan Keuangan Kas RT & Sosial")
        st.info("Fitur pembukuan kas aktif.")

    elif menu == "🖨️ Cetak Rekap PDF":
        if st.button("⬅️ Kembali ke Beranda", key="back_from_pdf_v10"):
            st.session_state.selected_menu = "Beranda / Dashboard"
            st.rerun()
        st.subheader("🖨️ Cetak Rekapitulasi PDF")
        st.success("Menu cetak PDF siap digunakan.")
