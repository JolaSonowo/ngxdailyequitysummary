# 5. Download Section
st.divider()
st.subheader("📁 Export Reports")

btn_col1, btn_col2 = st.columns(2)

# --- EXCEL DOWNLOAD ---
if not gainers_df.empty or not losers_df.empty:
    excel_bio = io.BytesIO()
    # Create the Excel file with two sheets
    with pd.ExcelWriter(excel_bio, engine='openpyxl') as writer:
        if not gainers_df.empty:
            gainers_df.to_excel(writer, sheet_name='Top Gainers', index=False)
        if not losers_df.empty:
            losers_df.to_excel(writer, sheet_name='Top Losers', index=False)
    
    excel_bio.seek(0)
    
    with btn_col1:
        st.download_button(
            label="📊 Download Excel Report",
            data=excel_bio.getvalue(),
            file_name=f"NGX_Report_{datetime.date.today()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

# --- WORD DOWNLOAD ---
if not gainers_df.empty or not losers_df.empty:
    doc = Document()
    doc.add_heading(f'NGX Market Summary - {datetime.date.today()}', 0)
    
    # Add Gainers Table
    if not gainers_df.empty:
        doc.add_heading('Top Gainers', level=1)
        table = doc.add_table(rows=1, cols=len(gainers_df.columns))
        table.style = 'Table Grid'
        hdr_cells = table.rows[0].cells
        for i, column in enumerate(gainers_df.columns):
            hdr_cells[i].text = column
        for index, row in gainers_df.iterrows():
            row_cells = table.add_row().cells
            for i, value in enumerate(row):
                row_cells[i].text = str(value)

    # Add Losers Table
    if not losers_df.empty:
        doc.add_heading('Top Losers', level=1)
        table_l = doc.add_table(rows=1, cols=len(losers_df.columns))
        table_l.style = 'Table Grid'
        hdr_cells_l = table_l.rows[0].cells
        for i, column in enumerate(losers_df.columns):
            hdr_cells_l[i].text = column
        for index, row in losers_df.iterrows():
            row_cells = table_l.add_row().cells
            for i, value in enumerate(row):
                row_cells[i].text = str(value)

    word_bio = io.BytesIO()
    doc.save(word_bio)
    
    with btn_col2:
        st.download_button(
            label="📝 Download Word Report",
            data=word_bio.getvalue(),
            file_name=f"NGX_Report_{datetime.date.today()}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True
        )
