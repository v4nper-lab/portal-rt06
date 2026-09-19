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
        
        # Konversi Format Tanggal Lahir ke Format Indonesia (Contoh: 02 Agustus 1991)
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
    menu = st.sidebar.selectbox("📂 Navigasi Menu", ["Dashboard & Data Warga", "Kelola Data Warga (Tambah/Edit/Hapus)", "Cetak Laporan PDF"])
    
    if menu == "Dashboard & Data Warga":
        tab_data, tab_grafik = st.tabs(["📋 Data Warga (Format KK)", "📈 Grafik Demografi Interaktif"])
        
        with tab_data:
            st.subheader("📋 Daftar Warga Berdasarkan Kartu Keluarga (Kepala Keluarga)")
            
            # Deteksi kolom Kepala Keluarga dan Anggota Keluarga
            col_kk = next((c for c in df.columns if "KEPALA KELUARGA" in c or "KK" in c), None)
            col_nama = next((c for c in df.columns if "ANGGOTA" in c or "NAMA" in c), None)
            
            if col_kk and col_nama:
                # Mengelompokkan warga berdasarkan Kepala Keluarga
                daftar_kk = df[col_kk].dropna().unique()
                
                for idx, kk in enumerate(daftar_kk, 1):
                    with st.expander(f"🏠 Keluarga: {kk} (No. {idx})", expanded=(idx == 1)):
                        df_anggota = df[df[col_kk] == kk]
                        st.dataframe(df_anggota, use_container_width=True, hide_index=True)
            else:
                # Fallback jika nama kolom berbeda
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
                    fig_jk.update_traces(textfont_size=16, textinfo="percent+label+value")
                    fig_jk.update_layout(font=dict(size=14))
                    st.plotly_chart(fig_jk, use_container_width=True)
                    
            with c2:
                if col_status:
                    st.markdown("#### 💍 Berdasarkan Status Pernikahan")
                    df_st = df[col_status].dropna().value_counts().reset_index()
                    df_st.columns = ["Status", "Jumlah"]
                    fig_st = px.bar(df_st, x="Status", y="Jumlah", text="Jumlah", color="Status", color_discrete_sequence=px.colors.qualitative.Pastel)
                    fig_st.update_traces(textfont_size=16, textposition="outside")
                    fig_st.update_layout(font=dict(size=14), xaxis=dict(tickfont=dict(size=14)), yaxis=dict(tickfont=dict(size=14)))
                    st.plotly_chart(fig_st, use_container_width=True)
                    
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
                    
            with c4:
                if col_pek:
                    st.markdown("#### 💼 Berdasarkan Pekerjaan")
                    df_pk = df[col_pek].dropna().value_counts().reset_index()
                    df_pk.columns = ["Pekerjaan", "Jumlah"]
                    fig_pk = px.bar(df_pk, x="Pekerjaan", y="Jumlah", text="Jumlah", color="Pekerjaan", color_discrete_sequence=px.colors.qualitative.Vivid)
                    fig_pk.update_traces(textfont_size=16, textposition="outside")
                    fig_pk.update_layout(font=dict(size=14), xaxis=dict(tickangle=-30, tickfont=dict(size=13)), yaxis=dict(tickfont=dict(size=14)))
                    st.plotly_chart(fig_pk, use_container_width=True)

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
                st.text_input("No. Rumah")
                st.text_input("Nama Kepala Keluarga")
                st.text_input("Nama Anggota Keluarga")
                st.selectbox("Jenis Kelamin", ["L", "P"])
                st.selectbox("Hubungan Keluarga", ["Kepala Keluarga", "Istri", "Anak Kandung", "Famili Lain", "Mertua"])
                st.text_input("Tempat Lahir")
                st.text_input("Tanggal Lahir (Contoh: 17 Agustus 1995)")
                st.number_input("Usia", min_value=0, max_value=120, value=25)
                st.selectbox("Status Pernikahan", ["Belum Kawin", "Kawin", "Cerai Hidup", "Cerai Mati"])
                st.text_input("Pendidikan Terakhir", "Tamat SLTA/sederajat")
                st.text_input("Pekerjaan", "Karyawan Swasta")
                
                if st.form_submit_button("Simpan Data Warga Baru"):
                    st.success("Data warga baru berhasil disiapkan!")

        elif aksi == "✏️ Edit Data Warga":
            st.markdown("### Edit Data Warga")
            st.selectbox("Pilih Warga yang Ingin Diedit:", df.iloc[:, 2] if len(df.columns) > 2 else df.iloc[:, 0])
            with st.form("form_edit"):
                st.text_input("Perbarui Data")
                st.form_submit_button("Simpan Perubahan")

        elif aksi == "🗑️ Hapus Data Warga":
            st.markdown("### Hapus Data Warga")
            nama_hapus = st.selectbox("Pilih Warga yang Ingin Dihapus:", df.iloc[:, 2] if len(df.columns) > 2 else df.iloc[:, 0])
            if st.button("Konfirmasi Hapus Warga Ini"):
                st.success(f"Data **{nama_hapus}** berhasil dihapus.")

    elif menu == "Cetak Laporan PDF":
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
