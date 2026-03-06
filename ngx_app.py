import streamlit as st
import datetime
import requests
import pandas as pd
import io
import pytz
from docx import Document

# 1. Page Configuration
st.set_page_config(page_title="NGX Market Dashboard", layout="wide")

# --- TIMEZONE LOGIC ---
def get_wat_time():
    utc_now = datetime.datetime.now(datetime.timezone.utc)
    wat_tz = pytz.timezone('Africa/Lagos')
    return utc_now.astimezone(wat_tz)

# 2. DATA FETCHING (Strictly limited to 7)
@st.cache_data(ttl=60)
def get_ngx_api_data(endpoint):
    url = f"https://doclib.ngxgroup.com/REST/api/mrkstat/{endpoint}"
    headers = {"User-Agent": "Mozilla/5.0", "Referer": "https://ngxgroup.com/"}
    try:
        response = requests.get(url, headers=headers, timeout=10)
        data = response.json()
        
        # STAGE 1: Slice the raw list to 7 items immediately
        limited_data = data[:7]
        
        output = []
        for item in limited_data:
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

# 3. HEADER
current_wat = get_wat_time()
st.title("NGX Daily Equity Summary")
st.write(f"Last Updated: {current_wat.strftime('%I:%M:%S %p')} WAT")

# 4. MAIN DASHBOARD CONTENT
col1, col2 = st.columns(2)

# Fetch data
gainers_df = get_ngx_api_data("topsymbols")
losers_df = get_ngx_api_data("bottomsymbols")

with col1:
    st.success("**Top 7 Advancers**")
    if not gainers_df.empty:
        # STAGE 2: Force display limit to 7
        st.dataframe(gainers_df.head(7), use_container_width=True, hide_index=True)
    else:
        st.warning("No data.")

with col2:
    st.error("**Top 7 Decliners**")
    if not losers_df.empty:
        # STAGE 2: Force display limit to 7
        st.dataframe(losers_df.head(7), use_container_width=True, hide_index=True)
    else:
        st.warning("No data.")

# 5. DOWNLOAD SECTION (Ensuring reports also only have 7)
st.divider()
st.subheader("Export Reports (Top 7 Only)")

# Final check: Truncate dataframes to 7 before generating files
gainers_final = gainers_df.head(7)
losers_final = losers_df.head(7)

btn_col1, btn_col2 = st.columns(2)

if not gainers_final.empty or not losers_final.empty:
    # Excel Logic
    excel_bio = io.BytesIO()
    with pd.ExcelWriter(excel_bio, engine='openpyxl') as writer:
        gainers_final.to_excel(writer, sheet_name='Top 7 Gainers', index=False)
        losers_final.to_excel(writer, sheet_name='Top 7 Losers', index=False)
    excel_bio.seek(0)

    # Word Logic
    doc = Document()
    doc.add_heading('NGX Top 7 Market Report', 0)
    
    for title, df in [("Top 7 Gainers", gainers_final), ("Top 7 Losers", losers_final)]:
        if not df.empty:
            doc.add_heading(title, level=1)
            table = doc.add_table(rows=1, cols=len(df.columns))
            table.style = 'Table Grid'
            for i, col in enumerate(df.columns):
                table.rows[0].cells[i].text = col
            for _, row in df.iterrows():
                cells = table.add_row().cells
                for i, val in enumerate(row):
                    cells[i].text = str(val)

    word_bio = io.BytesIO()
    doc.save(word_bio)

    with btn_col1:
        st.download_button("Download Excel", excel_bio.getvalue(), "NGX_7.xlsx", use_container_width=True)
    with btn_col2:
        st.download_button("Download Word", word_bio.getvalue(), "NGX_7.docx", use_container_width=True)            
            output.append({
                "Symbol": item.get('SYMBOL', 'N/A'),
                "Price": todays_close,
                "Change (N)": price_change,
                "% Change": round(pc_val, 2)
            })
        return pd.DataFrame(output)
    except Exception:
        return pd.DataFrame()

# 3. HEADER
current_wat = get_wat_time()
st.title("NGX Daily Equity Summary")
st.markdown(f"**{current_wat.strftime('%d %b %Y')}** | **{current_wat.strftime('%I:%M %p')} WAT**")

# 4. MAIN DASHBOARD CONTENT
col1, col2 = st.columns(2)

gainers_df = get_ngx_api_data("topsymbols")
losers_df = get_ngx_api_data("bottomsymbols")

with col1:
    st.success("**Top 7 Advancers**")
    if not gainers_df.empty:
        st.dataframe(gainers_df.head(7), use_container_width=True, hide_index=True)
    else:
        st.warning("No gainers data available currently.")

with col2:
    st.error("**Top 7 Decliners**")
    if not losers_df.empty:
        st.dataframe(losers_df.head(7), use_container_width=True, hide_index=True)
    else:
        st.warning("No losers data available currently.")

# 5. DOWNLOAD SECTION
st.divider()
st.subheader("Export Reports")

btn_col1, btn_col2 = st.columns(2)

# --- EXCEL DOWNLOAD LOGIC ---
if not gainers_df.empty or not losers_df.empty: 
    excel_bio = io.BytesIO()
    with pd.ExcelWriter(excel_bio, engine='openpyxl') as writer:
        if not gainers_df.empty:
            gainers_df.to_excel(writer, sheet_name='Top Gainers', index=False)
        if not losers_df.empty:
            losers_df.to_excel(writer, sheet_name='Top Losers', index=False)
    
    excel_bio.seek(0)
    
    with btn_col1:
        st.download_button(
            label="Download Excel Report",
            data=excel_bio.getvalue(),
            file_name=f"NGX_Report_{current_wat.strftime('%Y-%m-%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

# --- WORD DOWNLOAD LOGIC ---
if not gainers_df.empty or not losers_df.empty:
    doc = Document()
    doc.add_heading(f'NGX Market Summary', 0)
    doc.add_paragraph(f"Generated on: {current_wat.strftime('%d %b %Y at %I:%M:%S %p')} WAT")
    
    if not gainers_df.empty:
        doc.add_heading('Top 7 Gainers', level=1)
        table = doc.add_table(rows=1, cols=len(gainers_df.columns))
        table.style = 'Table Grid'
        for i, column in enumerate(gainers_df.columns):
            table.rows[0].cells[i].text = column
        for _, row in gainers_df.iterrows():
            row_cells = table.add_row().cells
            for i, value in enumerate(row):
                row_cells[i].text = str(value)

    if not losers_df.empty:
        doc.add_heading('Top 7 Losers', level=1)
        table_l = doc.add_table(rows=1, cols=len(losers_df.columns))
        table_l.style = 'Table Grid'
        for i, column in enumerate(losers_df.columns):
            table_l.rows[0].cells[i].text = column
        for _, row in losers_df.iterrows():
            row_cells = table_l.add_row().cells
            for i, value in enumerate(row):
                row_cells[i].text = str(value)

    word_bio = io.BytesIO()
    doc.save(word_bio)
    
    with btn_col2:
        st.download_button(
            label="Download Word Report",
            data=word_bio.getvalue(),
            file_name=f"NGX_Report_{current_wat.strftime('%Y-%m-%d')}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True
        )
