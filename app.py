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

# Custom CSS responsif untuk Mobile & Desktop, Ukuran Teks Menu Besar, dan Warna-Warni Keren
st.markdown("""
<style>
    .stApp {
        background: linear-gradient(135deg, #f1f5f9 0%, #cbd5e1 100%);
    }
    .metric-card {
        background: rgba(255, 255, 255, 0.9);
        backdrop-filter: blur(12px);
        border: 1px solid rgba(226, 232, 240, 0.9);
        padding: 20px;
        border-radius: 16px;
        box-shadow: 0 10px 25px -5px rgba(0, 0, 0, 0.05);
        margin-bottom: 12px;
    }
    /* Styling Tombol Menu Utama Menjadi Besar dan Berwarna-Warni Keren */
    .stButton button {
        font-size: 18px !important;
        font-weight: 700 !important;
        padding: 20px 24px !important;
        border-radius: 14px !important;
        border: none !important;
        color: white !important;
        box-shadow: 0 8px 15px rgba(0,0,0,0.1) !important;
        transition: all 0.3s ease !important;
        width: 100% !important;
    }
    .stButton button:hover {
        transform: translateY(-3px) !important;
        box-shadow: 0 12px 20px rgba(0,0,0,0.15) !important;
    }
    /* Variasi Warna-Warni Tombol Menu */
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

# Header Utama Portal RT 06 (Logo diperbesar 2x lipat tanpa background putih dengan efek CSS mix-blend-mode)
col_logo, col_title = st.columns([1, 4])
with col_logo:
    logo_path = "logo_rt06.jpg"
    if os.path.exists(logo_path):
        st.markdown('<div style="mix-blend-mode: multiply;">', unsafe_allow_html=True)
        st.image(logo_path, width=400)
        st.markdown('</div>', unsafe_allow_html=True)
    else:
        st.image("logo r6.jpg", width=400) if os.path.exists("logo r6.jpg") else st.write("🏠")

with col_title:
    st.markdown("<br>", unsafe_allow_html=True)
    st.title("🏠 PORTAL RT 06 / RW 14")
    st.markdown("**Griya Permata Raya** • Desa Nanjung Mekar, Rancaekek")

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
        
        col_kk_candi = [c for c in df.columns if "KEPALA" in c or "KK" in c]
        col_kk = col_kk_candi[0] if col_kk_candi else df.columns[2]
        
        col_rumah_candi = [c for c in df.columns if "RUMAH" in c or "ALAMAT" in c]
        col_rumah = col_rumah_candi[0] if col_rumah_candi else None

        if col_kk:
            df[col_kk] = df[col_kk].replace('', pd.NA)
            df[col_kk] = df[col_kk].ffill()
            
        if col_rumah:
            df[col_rumah] = df[col_rumah].replace('', pd.NA)
            df[col_rumah] = df[col_rumah].ffill()

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
        <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(150px, 1fr)); gap: 12px; margin-top: 10px; margin-bottom: 25px;">
            <div class="metric-card" style="border-left: 4px solid #2563eb;">
                <div style="font-size: 11px; text-transform: uppercase; font-weight: 600; color: #64748b;">Jumlah KK</div>
                <div style="font-size: 24px; font-weight: 800; color: #1e3a8a; margin-top: 4px;">{total_kk} <span style="font-size: 13px; font-weight: 500; color: #64748b;">KK</span></div>
            </div>
            <div class="metric-card" style="border-left: 4px solid #059669;">
                <div style="font-size: 11px; text-transform: uppercase; font-weight: 600; color: #64748b;">Total Jiwa</div>
                <div style="font-size: 24px; font-weight: 800; color: #065f46; margin-top: 4px;">{total_jiwa} <span style="font-size: 13px; font-weight: 500; color: #64748b;">Jiwa</span></div>
            </div>
            <div class="metric-card" style="border-left: 4px solid #0284c7;">
                <div style="font-size: 11px; text-transform: uppercase; font-weight: 600; color: #64748b;">Laki-laki</div>
                <div style="font-size: 24px; font-weight: 800; color: #0369a1; margin-top: 4px;">{jml_l}</div>
            </div>
            <div class="metric-card" style="border-left: 4px solid #db2777;">
                <div style="font-size: 11px; text-transform: uppercase; font-weight: 600; color: #64748b;">Perempuan</div>
                <div style="font-size: 24px; font-weight: 800; color: #9d174d; margin-top: 4px;">{jml_p}</div>
            </div>
            <div class="metric-card" style="border-left: 4px solid #d97706;">
                <div style="font-size: 11px; text-transform: uppercase; font-weight: 600; color: #64748b;">Balita (0-5 th)</div>
                <div style="font-size: 24px; font-weight: 800; color: #b45309; margin-top: 4px;">{jml_balita}</div>
            </div>
            <div class="metric-card" style="border-left: 4px solid #7c3aed;">
                <div style="font-size: 11px; text-transform: uppercase; font-weight: 600; color: #64748b;">Lansia (>60 th)</div>
                <div style="font-size: 24px; font-weight: 800; color: #5b21b6; margin-top: 4px;">{jml_lansia}</div>
            </div>
        </div>
        """, unsafe_allow_html=True)

        st.write("---")
        st.markdown("### 🚀 Menu Utama Portal RT 06")
        
        # Tombol Menu Ukuran Besar & Berwarna-Warni di Depan Dashboard
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
        
        daftar_kk = df[col_kk].dropna().astype(str).str.strip()
        daftar_kk = sorted(list(set([x for x in daftar_kk if x != "" and x.lower() != "nan" and x.lower() != "none"])))
        
        pilihan_kk = st.selectbox("Pilih Kepala Keluarga:", daftar_kk)
        
        if pilihan_kk:
            df_keluarga = df[df[col_kk].astype(str).str.strip().str.lower() == pilihan_kk.strip().lower()].copy()
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
                daftar_no_rumah = sorted(list(set(df[col_rumah].dropna().astype(str).tolist()))) if col_rumah else ["B3-01", "B3-02"]
                no_rumah = st.selectbox("No. Rumah", daftar_no_rumah)
                
                nama_kk = st.selectbox("Nama Kepala Keluarga", sorted(list(set(df[col_kk].dropna().astype(str).tolist()))))
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
        
        total_kk_rw = df[col_kk].nunique() if col_kk in df.columns else 0
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
                f"{df[col_rumah].nunique() if col_rumah in df.columns else 0} Rumah"
            ]
        }
        df_rekap_rw = pd.DataFrame(data_rekap_rw)
        st.dataframe(df_rekap_rw, use_container_width=True, hide_index=True)

    elif menu == "💰 Laporan Kas RT & Sosial (Perelek R6 Suayunan)":
        if st.button("⬅️ Kembali ke Beranda"):
            st.session_state.selected_menu = "Beranda / Dashboard"
            st.rerun()
        st.subheader("💰 Laporan Kas RT & Kas Sosial (Perelek R6 Suayunan)")
        st.markdown("Pusat transparansi keuangan warga RT 06, termasuk iuran kas RT serta pengelolaan dana sosial melalui **Perelek R6 Suayunan**.")
        
        tab_kas1, tab_kas2 = st.tabs(["📊 Kas RT 06", "🌾 Kas Sosial (Perelek R6 Suayunan)"])
        
        with tab_kas1:
            st.markdown("### Rekapitulasi Keuangan Kas RT 06")
            st.info("💡 Dana Kas RT digunakan untuk keperluan kebersihan lingkungan, keamanan, dan fasilitas umum warga.")
            
            # Tabel simulasi Kas RT
            data_kas_rt = {
                "No": [1, 2, 3, 4],
                "Keterangan Transaksi": ["Saldo Awal Bulan", "Iuran Warga Bulanan (Periode Berjalan)", "Pengeluaran Perbaikan Lampu Jalan", "Saldo Akhir Kas RT"],
                "Jenis": ["Masuk", "Masuk", "Keluar", "Total Saldo"],
                "Jumlah (Rp)": ["Rp 1.500.000", "Rp 2.400.000", "Rp 350.000", "Rp 3.550.000"]
            }
            st.dataframe(pd.DataFrame(data_kas_rt), use_container_width=True, hide_index=True)

        with tab_kas2:
            st.markdown("### Laporan Dana Sosial & Perelek R6 Suayunan")
            st.info("🌾 Program **Perelek R6 Suayunan** merupakan wujud gotong royong warga RT 06 untuk dana sosial kemasyarakatan (bantuan warga sakit, kedukaan, dll).")
            
            data_perelek = {
                "No": [1, 2, 3],
                "Uraian / Kegiatan Sosial": ["Saldo Kotak Sosial Perelek Sebelumnya", "Pemasukan Hasil Perelek Warga Bulanan", "Penyaluran Santunan Warga Sakit / Kedukaan"],
                "Status": ["Saldo", "Masuk", "Keluar"],
                "Nominal": ["Rp 750.000", "Rp 600.000", "Rp 250.000 (Saldo Akhir: Rp 1.100.000)"]
            }
            st.dataframe(pd.DataFrame(data_perelek), use_container_width=True, hide_index=True)

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
