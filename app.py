elif menu == "🗂️ Cetak Kartu Keluarga (KK)":
        if st.button("⬅️ Kembali ke Beranda", key="back_from_kk_v22"):
            st.session_state.selected_menu = "Beranda / Dashboard"
            st.rerun()
        st.subheader("🗂️ Cetak Kartu Keluarga (KK)")
        
        df_ffill = df.copy()
        if col_rumah:
            df_ffill[col_rumah] = df_ffill[col_rumah].replace('', pd.NA).ffill()
        if col_kk:
            df_ffill[col_kk] = df_ffill[col_kk].replace('', pd.NA).ffill()

        # Ambil daftar unik berdasarkan Nomor Rumah atau Kepala Keluarga tanpa spasi error
        list_kk_opsi = []
        for _, r_item in df_ffill.iterrows():
            kk_val = str(r_item.get(col_kk, "")).strip()
            rmh_val = str(r_item.get(col_rumah, "")).strip() if col_rumah else "-"
            if kk_val and kk_val.lower() not in ['', 'nan', 'none']:
                display_str = f"Rumah {rmh_val} - KK: {kk_val}"
                if display_str not in list_kk_opsi:
                    list_kk_opsi.append(display_str)

        pilihan_kk_str = st.selectbox("Pilih Rumah / Kepala Keluarga:", sorted(list_kk_opsi), key="select_kk_main_v22")
        
        if pilihan_kk_str:
            parts = pilihan_kk_str.split(" - KK: ")
            rmh_target = parts[0].replace("Rumah ", "").strip()
            kk_target = parts[1].strip()

            if col_rumah:
                df_ffill['_RUMAH_FILLED'] = df_ffill[col_rumah].ffill()
                df_keluarga = df_ffill[df_ffill['_RUMAH_FILLED'].astype(str).str.strip().str.lower() == rmh_target.lower()].copy()
                df_keluarga = df_keluarga.drop(columns=['_RUMAH_FILLED'])
            else:
                df_keluarga = df_ffill[df_ffill[col_kk].astype(str).str.strip().str.lower() == kk_target.lower()].copy()

            cols_tampilan_web = [c for c in df_keluarga.columns if c != col_kk and "URUT" not in c and c != "NO" and not c.startswith('_')]
            
            st.markdown(f"**No. Rumah:** {rmh_target} | **Kepala Keluarga:** {kk_target}")
            st.dataframe(df_keluarga[cols_tampilan_web], use_container_width=True, hide_index=True)

            st.markdown("---")
            if st.button("🖨️ Cetak / Download Kartu Keluarga (PDF)", key="btn_download_pdf_kk_v2"):
                try:
                    pdf_buffer = io.BytesIO()
                    doc = SimpleDocTemplate(pdf_buffer, pagesize=landscape(letter), rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=30)
                    elements = []
                    styles = getSampleStyleSheet()
                    
                    title_style = ParagraphStyle(
                        'TitleStyle',
                        parent=styles['Heading1'],
                        fontName='Helvetica-Bold',
                        fontSize=16,
                        textColor=colors.HexColor("#1e3a8a"),
                        alignment=1,
                        spaceAfter=15
                    )
                    
                    elements.append(Paragraph("KARTU KELUARGA (KK) - RT 06 / RW 14", title_style))
                    elements.append(Paragraph(f"<b>No. Rumah:</b> {rmh_target} &nbsp;&nbsp;&nbsp;&nbsp; <b>Kepala Keluarga:</b> {kk_target}", styles['Normal']))
                    elements.append(Spacer(1, 15))
                    
                    table_data = [cols_tampilan_web]
                    for _, r_row in df_keluarga[cols_tampilan_web].iterrows():
                        table_data.append([str(r_row.get(c, '')) for c in cols_tampilan_web])
                    
                    t_pdf = Table(table_data, repeatRows=1)
                    t_pdf.setStyle(TableStyle([
                        ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#2563eb")),
                        ('TEXTCOLOR', (0,0), (-1,0), colors.white),
                        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
                        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
                        ('FONTSIZE', (0,0), (-1,0), 9),
                        ('BOTTOMPADDING', (0,0), (-1,0), 8),
                        ('BACKGROUND', (0,1), (-1,-1), colors.HexColor("#f8fafc")),
                        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#cbd5e1")),
                        ('FONTNAME', (0,1), (-1,-1), 'Helvetica'),
                        ('FONTSIZE', (0,1), (-1,-1), 8),
                    ]))
                    elements.append(t_pdf)
                    doc.build(elements)
                    
                    pdf_data = pdf_buffer.getvalue()
                    st.download_button(
                        label="📥 Klik di Sini untuk Mengunduh File PDF KK",
                        data=pdf_data,
                        file_name=f"KK_Rumah_{rmh_target}.pdf",
                        mime="application/pdf",
                        key="dl_pdf_file_kk_action_v2"
                    )
                except Exception as e:
                    st.error(f"Gagal membuat PDF: {e}")
