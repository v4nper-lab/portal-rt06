import streamlit as st
import pandas as pd
import plotly.express as px
import os
import io
import time
from datetime import datetime, date
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
        st.write("🏠")

with col_title:
    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown('<div class="jumbo-title">🏠 PORTAL RT 06 / RW 14</div>', unsafe_allow_html=True)
    st.markdown('<div class="jumbo-subtitle">Griya Permata Raya • Desa Nanjung Mekar, Rancaekek</div>', unsafe_allow_html=True)

st.write("---")

# FUNGSI MEMBACA EXCEL TANPA CACHE (Membaca file fisik langsung agar data selalu real-time dan akurat)
def load_data_rt06_direct():
    if not os.path.exists(FILE_EXCEL_WARGA):
        return pd.DataFrame()
    
    try:
        df = pd.read_excel(FILE_EXCEL_WARGA, header=3)
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
            
        # Pengurutan otomatis berdasarkan No. Rumah secara rapi
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

df = load_data_rt06_direct()

if not df.empty:
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
        "🛠️ Kelola Warga", 
        "📊 Rekapitulasi Administrasi RW",
        "💰 Laporan Kas RT & Sosial (Perelek R6 Sauyunan)",
        "🖨️ Cetak Rekap PDF"
    ]

    def update_menu_pilihan():
        st.session_state.selected_menu = st.session_state.widget_nav_selectbox

    selected_sidebar = st.sidebar.selectbox(
        "Pilih Halaman:", 
        daftar_menu_pilihan, 
        index=daftar_menu_pilihan.index(st.session_state.selected_menu) if st.session_state.selected_menu in daftar_menu_pilihan else 0,
        key="widget_nav_selectbox",
        on_change=update_menu_pilihan
    )

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
            if st.button("📋 Data Seluruh Warga", use_container_width=True, key="btn_m1"):
                st.session_state.selected_menu = "📋 Data Seluruh Warga"
                st.rerun()
            if st.button("🗂️ Cetak Kartu Keluarga (KK)", use_container_width=True, key="btn_m2"):
                st.session_state.selected_menu = "🗂️ Cetak Kartu Keluarga (KK)"
                st.rerun()
            if st.button("📊 Rekapitulasi Administrasi RW", use_container_width=True, key="btn_m3"):
                st.session_state.selected_menu = "📊 Rekapitulasi Administrasi RW"
                st.rerun()
            if st.button("💰 Laporan Kas RT & Sosial (Perelek R6 Sauyunan)", use_container_width=True, key="btn_m4"):
                st.session_state.selected_menu = "💰 Laporan Kas RT & Sosial (Perelek R6 Sauyunan)"
                st.rerun()
        with col_m2:
            if st.button("📈 Grafik Demografi", use_container_width=True, key="btn_m5"):
                st.session_state.selected_menu = "📈 Grafik Demografi"
                st.rerun()
            if st.button("🛠️ Kelola Warga", use_container_width=True, key="btn_m6"):
                st.session_state.selected_menu = "🛠️ Kelola Warga"
                st.rerun()
            if st.button("🖨️ Cetak Laporan Rekap PDF", use_container_width=True, key="btn_m7"):
                st.session_state.selected_menu = "🖨️ Cetak Rekap PDF"
                st.rerun()

        for _ in range(5):
            waktu_sekarang = datetime.now(ZoneInfo("Asia/Jakarta"))
            bulan_indo_nama = {1: "Januari", 2: "Februari", 3: "Maret", 4: "April", 5: "Mei", 6: "Juni", 7: "Juli", 8: "Agustus", 9: "September", 10: "Oktober", 11: "November", 12: "Desember"}
            tgl_str = f"{waktu_sekarang.day:02d} {bulan_indo_nama.get(waktu_sekarang.month, '')} {waktu_sekarang.year}"
            jam_str = waktu_sekarang.strftime("%H:%M:%S")
            
            placeholder_waktu.markdown(f"""
            <div style="background: rgba(255, 255, 255, 0.9); border: 1px solid #cbd5e1; padding: 8px 12px; border-radius: 10px; text-align: right;">
                <span style="font-size: 10px; color: #64748b;">🕒 Live Update (WIB):</span><br>
                <strong style="font-size: 12px; color: #0f172a;">{tgl_str} | {jam_str} WIB</strong>
            </div>
            """, unsafe_allow_html=True)
            time.sleep(1)
        st.rerun()

    elif menu == "📋 Data Seluruh Warga":
        if st.button("⬅️ Kembali ke Beranda", key="back_warga"):
            st.session_state.selected_menu = "Beranda / Dashboard"
            st.rerun()
        st.subheader("📋 Data Keseluruhan Warga (Real-Time Database)")
        
        st.markdown("💡 **Mode Interaktif Excel:** Setiap perubahan yang Anda simpan akan langsung memperbarui file database fisik secara permanen. Warga dengan status domisili **luar NM** otomatis diberi **warna latar kuning lembut**.")

        def highlight_luar_nm(row):
            row_str = str(row.values).lower()
            if "luar nm" in row_str:
                return ['background-color: #fef08a'] * len(row)
            return [''] * len(row)

        try:
            styled_df = df.style.apply(highlight_luar_nm, axis=1)
            edited_df = st.data_editor(
                styled_df,
                num_rows="dynamic",
                use_container_width=True,
                key="editor_tabel_warga_interaktif"
            )
        except:
            edited_df = st.data_editor(
                df,
                num_rows="dynamic",
                use_container_width=True,
                key="editor_tabel_warga_interaktif_fallback"
            )

        if st.button("💾 Simpan Perubahan ke Database Excel", key="save_warga_db"):
            try:
                if isinstance(edited_df, pd.DataFrame):
                    col_rumah_save = next((col for col in edited_df.columns if "RUMAH" in col or "ALAMAT" in col), None)
                    col_kk_save = next((col for col in edited_df.columns if "KEPALA" in col or "KK" in col), None)
                    
                    if col_rumah_save:
                        edited_df['_TEMP_RUMAH'] = edited_df[col_rumah_save].ffill()
                        if col_kk_save:
                            edited_df['_TEMP_KK'] = edited_df[col_kk_save].ffill()
                            edited_df = edited_df.sort_values(by=['_TEMP_RUMAH', '_TEMP_KK'], key=lambda x: x.astype(str).str.lower()).reset_index(drop=True)
                        else:
                            edited_df = edited_df.sort_values(by='_TEMP_RUMAH', key=lambda x: x.astype(str).str.lower()).reset_index(drop=True)
                        edited_df = edited_df.drop(columns=[c for c in edited_df.columns if c.startswith('_TEMP_')])

                    import openpyxl
                    wb = openpyxl.Workbook()
                    ws = wb.active
                    ws.title = "Data Warga"
                    ws.append(["DATA WARGA RT 06 RW 14"])
                    ws.append([])
                    ws.append([])
                    ws.append(list(edited_df.columns))
                    for _, r in edited_df.iterrows():
                        ws.append(list(r.values))
                    wb.save(FILE_EXCEL_WARGA)

                    st.success("✅ Perubahan data warga berhasil disimpan permanen ke file Excel!")
                    time.sleep(1)
                    st.rerun()
            except Exception as e:
                st.error(f"❌ Gagal menyimpan perubahan: {e}")

    elif menu == "🗂️ Cetak Kartu Keluarga (KK)":
        if st.button("⬅️ Kembali ke Beranda", key="back_kk"):
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
        
        pilihan_kk = st.selectbox("Pilih Kepala Keluarga:", daftar_kk, key="select_kk_print")
        
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
                type="primary",
                key="dl_pdf_kk_btn"
            )

    elif menu == "📈 Grafik Demografi":
        if st.button("⬅️ Kembali ke Beranda", key="back_grafik"):
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
            st.plotly_chart(fig_jk, use_container_width=True, key="chart_jk_pie")

        if col_status:
            st.markdown("---")
            df_st = df[col_status].dropna().value_counts().reset_index()
            df_st.columns = ["Status", "Jumlah"]
            fig_st = px.bar(df_st, x="Status", y="Jumlah", text="Jumlah", title="💍 Distribusi Status Pernikahan Warga", color="Status", color_discrete_sequence=px.colors.qualitative.Pastel)
            fig_st.update_traces(textfont_size=16, textposition="outside")
            fig_st.update_layout(font=chart_font, title_font=title_font)
            st.plotly_chart(fig_st, use_container_width=True, key="chart_status_bar")

        if col_pend:
            st.markdown("---")
            df_pd = df[col_pend].dropna().value_counts().reset_index()
            df_pd.columns = ["Pendidikan", "Jumlah"]
            fig_pd = px.bar(df_pd, x="Pendidikan", y="Jumlah", text="Jumlah", title="🎓 Tingkat Pendidikan Terakhir Warga", color="Pendidikan", color_discrete_sequence=px.colors.qualitative.Vivid)
            fig_pd.update_traces(textfont_size=16, textposition="outside")
            fig_pd.update_layout(font=chart_font, title_font=title_font)
            st.plotly_chart(fig_pd, use_container_width=True, key="chart_pend_bar")

        if col_pek:
            st.markdown("---")
            df_pk = df[col_pek].dropna().value_counts().reset_index()
            df_pk.columns = ["Pekerjaan", "Jumlah"]
            fig_pk = px.bar(df_pk, x="Pekerjaan", y="Jumlah", text="Jumlah", title="💼 Distribusi Mata Pencaharian / Pekerjaan Warga", color="Pekerjaan", color_discrete_sequence=px.colors.qualitative.Safe)
            fig_pk.update_traces(textfont_size=16, textposition="outside")
            fig_pk.update_layout(font=chart_font, title_font=title_font)
            st.plotly_chart(fig_pk, use_container_width=True, key="chart_pek_bar")

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
            st.plotly_chart(fig_usia, use_container_width=True, key="chart_usia_pie")

    elif menu == "🛠️ Kelola Warga":
        if st.button("⬅️ Kembali ke Beranda", key="back_kelola"):
            st.session_state.selected_menu = "Beranda / Dashboard"
            st.rerun()
        st.subheader("🛠️ Kelola Data Warga (Input Warga Baru & Anggota Keluarga)")

        st.markdown("💡 **Formulir Isian Real-Time:** Pilih No. Rumah, masukkan Nama Kepala Keluarga & Anggota Keluarga. Data baru akan otomatis tersimpan langsung ke file fisik Excel tanpa tertimpa memori lama.")

        daftar_blok_lengkap = [
            "B3-01", "B3-02", "B3-03", "B3-04", "B3-05", "B3-06", "B3-07", "B3-08", "B3-09", "B3-10",
            "B3-11", "B3-12", "B3-13", "B3-14", "B3-15", "B3-16", "B3-17", "B3-18", "B3-19", "B3-20",
            "B4-01", "B4-02", "B4-03", "B4-04", "B4-05", "B4-06", "B4-07", "B4-08", "B4-09", "B4-10"
        ]
        if not df.empty and col_rumah:
            df_temp_blok = df.copy()
            df_temp_blok[col_rumah] = df_temp_blok[col_rumah].ffill()
            blok_dari_data = sorted(list(set(df_temp_blok[col_rumah].dropna().astype(str).tolist())))
            for b in blok_dari_data:
                if b not in daftar_blok_lengkap:
                    daftar_blok_lengkap.append(b)

        with st.form("form_tambah_warga_perfect_v19", clear_on_submit=False):
            st.markdown("#### 📝 Form Input Isian Terstruktur:")
            
            col_f1, col_f2 = st.columns(2)
            tgl_lhr, bln_lhr, thn_lhr = 1, 1, 1995

            # Kolom Kiri
            with col_f1:
                in_no_rumah = st.selectbox("No. Rumah", sorted(daftar_blok_lengkap), key="in_no_rmh_v19")
                in_nama_kk = st.text_input("Nama Kepala Keluarga", key="in_nama_kk_v19")
                in_nama_anggota = st.text_input("Nama Lengkap Anggota Keluarga", key="in_nama_anggota_v19")
                in_jk = st.selectbox("L/P", ["L", "P"], key="in_jk_v19")
                in_hub = st.selectbox("Hubungan Keluarga", ["Kepala Keluarga", "Istri", "Anak Kandung", "Famili Lain", "Mertua"], key="in_hub_v19")
                in_tmplhr = st.text_input("Tempat Lahir", key="in_tmplhr_v19")
                
                st.markdown("📅 **Tanggal Lahir:**")
                c_t1, c_t2, c_t3 = st.columns(3)
                with c_t1: tgl_lhr = st.number_input("Tgl", min_value=1, max_value=31, value=1, key="in_tgl_v19")
                with c_t2: bln_lhr = st.number_input("Bln", min_value=1, max_value=12, value=8, key="in_bln_v19")
                with c_t3: thn_lhr = st.number_input("Thn", min_value=1900, max_value=2026, value=1995, key="in_thn_v19")
                
                bulan_indo_nama = {1: "Januari", 2: "Februari", 3: "Maret", 4: "April", 5: "Mei", 6: "Juni", 7: "Juli", 8: "Agustus", 9: "September", 10: "Oktober", 11: "November", 12: "Desember"}
                in_tgllhr = f"{int(tgl_lhr):02d} {bulan_indo_nama.get(int(bln_lhr), '')} {int(thn_lhr)}"

            # Kolom Kanan
            with col_f2:
                usia_otomatis = 2026 - int(thn_lhr)
                if usia_otomatis < 0: usia_otomatis = 0
                in_usia = st.number_input("Usia (Otomatis)", min_value=0, max_value=120, value=int(usia_otomatis), key="in_usia_auto_v19")
                
                in_status_nikah = st.selectbox("Status (Status Perkawinan)", ["Belum Kawin", "Kawin", "Cerai Hidup", "Cerai Mati"], key="in_status_nikah_v19")
                in_agama = st.selectbox("Agama", ["Islam", "Kristen", "Katolik", "Hindu", "Buddha", "Konghucu"], key="in_agama_v19")
                in_goldarah = st.selectbox("Gol. Darah", ["A", "B", "AB", "O", "Tidak Tahu"], key="in_goldarah_v19")
                in_suku = st.selectbox("Etnis / Suku", ["Sunda", "Jawa", "Padang", "Batak", "Betawi", "Lainnya"], key="in_suku_v19")
                in_pend = st.selectbox("Pendidikan", ["Tamat SLTA/sederajat", "Tamat SLTP/sederajat", "Tamat SD/sederajat", "Diploma IV / Strata I", "Sedang SD/sedajerat", "Sedang SLTP/sederajat", "Belum / Tidak Sekolah"], key="in_pend_v19")
                in_pek = st.selectbox("Pekerjaan", ["Karyawan Swasta", "Wiraswasta", "Mengurus Rumah Tangga", "Belum Bekerja", "Pelajar", "PNS / TNI / Polri"], key="in_pek_v19")
                
                in_status_rumah = st.selectbox("Status Rumah (Stus Rumah)", ["Milik / Tetap", "Sewa/Kontrak", "Kosong"], key="in_status_rumah_v19")
                in_status_domisili = st.selectbox("Status Domisili", ["Nanjung Mekar", "luar NM"], key="in_status_domisili_v19")

            if st.form_submit_button("💾 Simpan Data ke Database Warga"):
                try:
                    df_raw_excel = pd.read_excel(FILE_EXCEL_WARGA, header=3)
                    df_raw_excel.columns = df_raw_excel.columns.astype(str).str.strip().str.upper()
                    
                    baris_baru_dict = {}
                    for col_excel in df_raw_excel.columns:
                        c_up = str(col_excel).upper()
                        if col_excel == "NO" or col_excel == "NO.":
                            continue
                        elif ("RUMAH" in c_up and "STATUS" not in c_up and "STUS" not in c_up) or "ALAMAT" in c_up:
                            if in_hub.lower() != "kepala keluarga":
                                baris_baru_dict[col_excel] = None
                            else:
                                baris_baru_dict[col_excel] = in_no_rumah
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
                            baris_baru_dict[col_excel] = in_tgllhr
                        elif "USIA" in c_up or "UMUR" in c_up:
                            baris_baru_dict[col_excel] = 2026 - int(thn_lhr)
                        elif c_up == "STATUS" or "KAWIN" in c_up or "NIKAH" in c_up or "PERKAWINAN" in c_up:
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
                        elif "STATUS RUMAH" in c_up or "STUS" in c_up:
                            baris_baru_dict[col_excel] = in_status_rumah
                        elif "STATUS DOMISILI" in c_up or ("STATUS" in c_up and "DOMISILI" in c_up):
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

                    # Urutkan otomatis berdasarkan blok rumah
                    col_rumah_save_db = next((col for col in df_updated_excel.columns if "RUMAH" in col or "ALAMAT" in col), None)
                    col_kk_save_db = next((col for col in df_updated_excel.columns if "KEPALA" in col or "KK" in col), None)

                    if col_rumah_save_db:
                        df_updated_excel['_TEMP_RUMAH'] = df_updated_excel[col_rumah_save_db].ffill()
                        if col_kk_save_db:
                            df_updated_excel['_TEMP_KK'] = df_updated_excel[col_kk_save_db].ffill()
                            df_updated_excel = df_updated_excel.sort_values(by=['_TEMP_RUMAH', '_TEMP_KK'], key=lambda x: x.astype(str).str.lower()).reset_index(drop=True)
                        else:
                            df_updated_excel = df_updated_excel.sort_values(by='_TEMP_RUMAH', key=lambda x: x.astype(str).str.lower()).reset_index(drop=True)
                        df_updated_excel = df_updated_excel.drop(columns=[c for c in df_updated_excel.columns if c.startswith('_TEMP_')])

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

                    st.success("✅ Data warga baru berhasil disimpan secara real-time ke file Excel!")
                    time.sleep(1)
                    st.rerun()
                except Exception as e:
                    st.error(f"❌ Gagal menyimpan data: {e}")

    elif menu == "📊 Rekapitulasi Administrasi RW":
        if st.button("⬅️ Kembali ke Beranda", key="back_rw"):
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

    elif menu == "💰 Laporan Kas RT & Sosial (Perelek R6 Sauyunan)":
        if st.button("⬅️ Kembali ke Beranda", key="back_kas"):
            st.session_state.selected_menu = "Beranda / Dashboard"
            st.rerun()
        st.subheader("💰 Input & Rekapitulasi Laporan Keuangan Kas RT & Kas Sosial (Perelek R6 Sauyunan)")
        st.markdown("💡 **Sistem Input Formulir Akuntansi:** Gunakan formulir di bawah ini untuk menambahkan transaksi baru secara akurat. Data dijamin tersimpan permanen dan langsung menghasilkan rekapan akuntansi yang benar.")

        def render_buku_kas_formulir(state_key, file_path, judul_buku, file_pdf_name, judul_pdf, pakai_logo=False):
            st.markdown(f"### 📝 Tambah Transaksi Baru: {judul_buku}")
            
            if st.button(f"🧹 Bersihkan Semua Data (Mulai dari Awal) - {judul_buku}", key=f"reset_{state_key}"):
                empty_df = pd.DataFrame(columns=["Tanggal", "Uraian / Keterangan Transaksi", "Debet (Masuk)", "Kredit (Keluar)"])
                st.session_state[state_key] = empty_df
                simpan_data_kas(file_path, empty_df)
                st.success(f"✅ Data {judul_buku} berhasil dikosongkan. Silakan mulai input dari awal!")
                st.rerun()

            with st.form(f"form_tambah_{state_key}", clear_on_submit=True):
                col_f1, col_f2 = st.columns(2)
                with col_f1:
                    tgl_input = st.text_input("Tanggal Transaksi (Contoh: 01/06/2026)", value=datetime.now(ZoneInfo("Asia/Jakarta")).strftime("%d/%m/%Y"), key=f"tgl_kas_{state_key}")
                    uraian_input = st.text_input("Uraian / Keterangan Transaksi", key=f"uraian_kas_{state_key}")
                with col_f2:
                    debet_input = st.number_input("Debet / Masuk (Rp)", min_value=0.0, step=1000.0, value=0.0, key=f"debet_kas_{state_key}")
                    kredit_input = st.number_input("Kredit / Keluar (Rp)", min_value=0.0, step=1000.0, value=0.0, key=f"kredit_kas_{state_key}")
                
                submitted = st.form_submit_button("➕ Tambahkan ke Pembukuan")
                if submitted:
                    if not uraian_input.strip():
                        st.warning("⚠️ Uraian / Keterangan transaksi tidak boleh kosong!")
                    else:
                        df_cur = st.session_state[state_key].copy()
                        baris_baru = pd.DataFrame({
                            "Tanggal": [tgl_input],
                            "Uraian / Keterangan Transaksi": [uraian_input],
                            "Debet (Masuk)": [float(debet_input)],
                            "Kredit (Keluar)": [float(kredit_input)]
                        })
                        df_updated = pd.concat([df_cur, baris_baru], ignore_index=True)
                        st.session_state[state_key] = df_updated
                        simpan_data_kas(file_path, df_updated)
                        st.success(f"✅ Transaksi '**{uraian_input}**' berhasil dicatat ke pembukuan!")
                        st.rerun()

            st.markdown(f"### 📊 Tabel Rekapitulasi Akuntansi: {judul_buku}")
            df_sumber = st.session_state[state_key].copy()
            df_tampil_live = hitung_dan_tampilkan_tabel_tunggal(df_sumber)

            if df_tampil_live.empty:
                st.info("ℹ️ Belum ada transaksi tercatat. Silakan tambahkan transaksi melalui formulir di atas.")
            else:
                col_aksi1, col_aksi2 = st.columns([2, 2])
                with col_aksi1:
                    if st.button(f"🗑️ Hapus Baris Transaksi Terakhir ({judul_buku})", key=f"del_{state_key}"):
                        if len(df_sumber) > 0:
                            df_sumber = df_sumber.iloc[:-1].reset_index(drop=True)
                            st.session_state[state_key] = df_sumber
                            simpan_data_kas(file_path, df_sumber)
                            st.success("✅ Baris transaksi terakhir berhasil dihapus!")
                            st.rerun()

                st.dataframe(df_tampil_live, use_container_width=True, hide_index=True)

            def format_rupiah_pdf(num):
                try:
                    n = float(num)
                    if n == 0: return "-"
                    return f"Rp {int(n):,}".replace(",", ".")
                except:
                    return "-"

            def buat_pdf_standar_akuntansi(df_lap, judul):
                buffer = io.BytesIO()
                doc = SimpleDocTemplate(buffer, pagesize=letter, rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=30)
                elements = []
                styles = getSampleStyleSheet()
                
                style_judul_pusat = ParagraphStyle('JudulPusat', parent=styles['Heading1'], fontSize=16, alignment=1, textColor=colors.HexColor('#1f2937'), fontName='Helvetica-Bold')
                style_subjudul_pusat = ParagraphStyle('SubJudulPusat', parent=styles['Normal'], fontSize=12, alignment=1, textColor=colors.HexColor('#4b5563'), fontName='Helvetica-Bold')
                
                if pakai_logo:
                    logo_file = "logo_perelek.png"
                    if not os.path.exists(logo_file):
                        logo_file = "logo_perelek.jpg"
                    
                    p_judul = Paragraph(f"<b>{judul}</b>", style_judul_pusat)
                    p_sub = Paragraph("<b>PERIODE TAHUN 2026</b>", style_subjudul_pusat)
                    
                    if os.path.exists(logo_file):
                        try:
                            img_logo = RLImage(logo_file, width=45, height=45)
                            t_header = Table([[img_logo, [p_judul, Spacer(1, 3), p_sub]]], colWidths=[55, 445])
                            t_header.setStyle(TableStyle([
                                ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                                ('ALIGN', (0,0), (0,0), 'CENTER'),
                                ('ALIGN', (1,0), (1,0), 'CENTER'),
                            ]))
                            elements.append(t_header)
                        except:
                            elements.append(p_judul)
                            elements.append(Spacer(1, 3))
                            elements.append(p_sub)
                    else:
                        elements.append(p_judul)
                        elements.append(Spacer(1, 3))
                        elements.append(p_sub)
                else:
                    elements.append(Paragraph(f"<b>{judul}</b>", style_judul_pusat))
                    elements.append(Spacer(1, 3))
                    elements.append(Paragraph("<b>PERIODE TAHUN 2026</b>", style_subjudul_pusat))
                
                elements.append(Spacer(1, 15))
                
                df_pdf_clean = hitung_dan_tampilkan_tabel_tunggal(df_lap)
                df_pdf_clean["Debet (Masuk)"] = df_pdf_clean["Debet (Masuk)"].apply(parsing_angka_aman).apply(format_Rupiah)
                df_pdf_clean["Kredit (Keluar)"] = df_pdf_clean["Kredit (Keluar)"].apply(parsing_angka_aman).apply(format_Rupiah)
                
                kolom = list(df_pdf_clean.columns)
                cell_s_left = ParagraphStyle('CellLeft', parent=styles['Normal'], fontSize=8.5, leading=10, alignment=0)
                cell_s_center = ParagraphStyle('CellCenter', parent=styles['Normal'], fontSize=8.5, leading=10, alignment=1)
                head_s = ParagraphStyle('Head', parent=styles['Normal'], fontSize=9, leading=11, textColor=colors.whitesmoke, fontName='Helvetica-Bold', alignment=1)
                
                t_data = [[Paragraph(c, head_s) for c in kolom]]
                for _, r in df_pdf_clean.iterrows():
                    row_cells = []
                    for col_name in kolom:
                        val_str = str(r[col_name]) if pd.notnull(r[col_name]) else ""
                        if "URAIAN" in col_name.upper():
                            row_cells.append(Paragraph(val_str, cell_s_left))
                        else:
                            row_cells.append(Paragraph(val_str, cell_s_center))
                    t_data.append(row_cells)
                    
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
                elements.append(Spacer(1, 15))
                
                waktu_pdf = datetime.now(ZoneInfo("Asia/Jakarta"))
                bulan_indo_nama = {1: "Januari", 2: "Februari", 3: "Maret", 4: "April", 5: "Mei", 6: "Juni", 7: "Juli", 8: "Agustus", 9: "September", 10: "Oktober", 11: "November", 12: "Desember"}
                tgl_cetak_pdf = f"Bandung, {waktu_pdf.day} {bulan_indo_nama.get(waktu_pdf.month, '')} {waktu_pdf.year}"
                
                ttd_data = [
                    [Paragraph("<b>Mengetahui,<br/>Ketua RT 06</b>", ParagraphStyle('T1', parent=styles['Normal'], alignment=1, fontSize=9)),
                     Paragraph(f"<b>{tgl_cetak_pdf}</b><br/>Bendahara RT 06", ParagraphStyle('T2', parent=styles['Normal'], alignment=1, fontSize=9))],
                    [Spacer(1, 35), Spacer(1, 35)],
                    [Paragraph("<b>( ......................................... )</b>", ParagraphStyle('T3', parent=styles['Normal'], alignment=1, fontSize=9)),
                     Paragraph("<b>( ......................................... )</b>", ParagraphStyle('T4', parent=styles['Normal'], alignment=1, fontSize=9))]
                ]
                t_ttd = Table(ttd_data, colWidths=[250, 250])
                elements.append(t_ttd)
                
                doc.build(elements)
                buffer.seek(0)
                return buffer.getvalue()

            pdf_bytes = buat_pdf_standar_akuntansi(st.session_state[state_key], judul_pdf)
            st.download_button(
                label=f"📥 Download PDF {judul_buku}",
                data=pdf_bytes,
                file_name=file_pdf_name,
                mime="application/pdf",
                type="primary",
                key=f"dl_pdf_kas_{state_key}"
            )

        tab_kas1, tab_kas2 = st.tabs(["📊 Buku Kas RT 06", "🌾 Buku Kas Sosial (Perelek)"])
        
        with tab_kas1:
            render_buku_kas_formulir('df_kas_rt_state', FILE_KAS_RT, "Buku Kas RT 06", "Laporan_Kas_RT06.pdf", "LAPORAN PERTANGGUNGJAWABAN KEUANGAN KAS RT 06", pakai_logo=False)

        with tab_kas2:
            render_buku_kas_formulir('df_kas_sosial_state', FILE_KAS_SOSIAL, "Buku Kas Sosial / Perelek", "Laporan_Kas_Sosial.pdf", "LAPORAN KAS PERELEK R6 SAUYUNAN", pakai_logo=True)

    elif menu == "🖨️ Cetak Rekap PDF":
        if st.button("⬅️ Kembali ke Beranda", key="back_rekap_pdf"):
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
            type="primary",
            key="dl_pdf_rekap_all"
        )

else:
    st.info("Pastikan file Excel data warga RT 06 sudah tersedia.")
