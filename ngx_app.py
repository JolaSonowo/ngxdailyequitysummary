import streamlit as st
import datetime
import requests
import pandas as pd
import io
import pytz
import time
from docx import Document

# 1. Page Configuration
st.set_page_config(page_title="NGX Market Dashboard", layout="wide")

# --- TIMEZONE LOGIC ---
def get_wat_time():
    """Returns the current time in West Africa Time (Lagos)"""
    utc_now = datetime.datetime.now(datetime.timezone.utc)
    wat_tz = pytz.timezone('Africa/Lagos')
    return utc_now.astimezone(wat_tz)

# 2. DATA FETCHING (Cached for 60 seconds)
@st.cache_data(ttl=60)
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

# 3. HEADER & LIVE CLOCK PLACEHOLDER
st.title("NGX Daily Equity Summary")

# We create an empty container that we will fill with the clock later
clock_placeholder = st.empty()

# 4. MAIN DASHBOARD CONTENT
col1, col2 = st.columns(2)

# Get the data (this is cached, so it's fast)
gainers_df = get_ngx_api_data("topsymbols")
losers_df = get_ngx_api_data("bottomsymbols")

with col1:
    st.success("###Top 5 Advancers")
    if not gainers_df.empty:
        st.dataframe(gainers_df, use_container_width=True, hide_index=True)
    else:
        st.warning("No gainers data available currently.")

with col2:
    st.error("###Top 5 Decliners")
    if not losers_df.empty:
        st.dataframe(losers_df, use_container_width=True, hide_index=True)
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
            file_name=f"NGX_Report_{get_wat_time().strftime('%Y-%m-%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

# --- WORD DOWNLOAD LOGIC ---
if not gainers_df.empty or not losers_df.empty:
    doc = Document()
    current_wat = get_wat_time()
    doc.add_heading(f'NGX Market Summary', 0)
    doc.add_paragraph(f"Generated on: {current_wat.strftime('%d %b %Y at %I:%M:%S %p')} WAT")
    
    if not gainers_df.empty:
        doc.add_heading('Top Gainers', level=1)
        table = doc.add_table(rows=1, cols=len(gainers_df.columns))
        table.style = 'Table Grid'
        for i, column in enumerate(gainers_df.columns):
            table.rows[0].cells[i].text = column
        for _, row in gainers_df.iterrows():
            row_cells = table.add_row().cells
            for i, value in enumerate(row):
                row_cells[i].text = str(value)

    if not losers_df.empty:
        doc.add_heading('Top Losers', level=1)
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

# 6. LIVE CLOCK (Runs at the very end)
# This loop updates only the placeholder we created at the top
while True:
    now = get_wat_time()
    clock_placeholder.markdown(
        f" **{now.strftime('%d %b %Y')}** | **{now.strftime('%I:%M:%S %p')} WAT**"
    )
    time.sleep(1)
