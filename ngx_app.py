import streamlit as st
import datetime
import requests
import pandas as pd
import io
from docx import Document

# 1. Page Configuration
st.set_page_config(page_title="NGX Market Dashboard", layout="wide")

# 2. Reuse your Data Logic
def get_ngx_api_data(endpoint):
    url = f"https://doclib.ngxgroup.com/REST/api/mrkstat/{endpoint}"
    headers = {"User-Agent": "Mozilla/5.0", "Referer": "https://ngxgroup.com/"}
    try:
        response = requests.get(url, headers=headers, timeout=10)
        data = response.json()
        output = []
        for item in data[:5]:
            last_close = float(item.get('LAST_CLOSE', 0))
            price_change = float(item.get('PERCENTAGE_CHANGE', 0))
            todays_close = float(item.get('TODAYS_CLOSE', 0))
            pc_val = (price_change / last_close * 100) if last_close != 0 else 0
            
            output.append({
                "Symbol": item.get('SYMBOL', 'N/A'),
                "Price": todays_close,
                "Change (N)": price_change,
                "% Change": round(pc_val, 2)
            })
        return pd.DataFrame(output)
    except:
        return pd.DataFrame()

# 3. Sidebar & Header
st.title("🇳🇬 NGX Live Market Dashboard")
st.write(f"Last updated: **{datetime.date.today().strftime('%d %b %Y')}**")

# 4. Main Content - Side by Side Tables
col1, col2 = st.columns(2)

with col1:
    st.success("### 📈 Top Gainers")
    gainers_df = get_ngx_api_data("topsymbols")
    if not gainers_df.empty:
        st.dataframe(gainers_df, use_container_width=True, hide_index=True)

with col2:
    st.error("### 📉 Top Losers")
    losers_df = get_ngx_api_data("bottomsymbols")
    if not losers_df.empty:
        st.dataframe(losers_df, use_container_width=True, hide_index=True)

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
