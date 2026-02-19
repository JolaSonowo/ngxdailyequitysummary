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
    utc_now = datetime.datetime.now(datetime.timezone.utc)
    wat_tz = pytz.timezone('Africa/Lagos')
    return utc_now.astimezone(wat_tz)

# 2. DATA FETCHING (Cached to prevent hitting the API every second)
@st.cache_data(ttl=60) # Updates data every 60 seconds
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

# 3. HEADER & LIVE CLOCK
st.title("🇳🇬 NGX Live Market Dashboard")

# Create a placeholder for the live clock
clock_placeholder = st.empty()

# 4. Main Content - Side by Side Tables
col1, col2 = st.columns(2)

# Load data once per run
gainers_df = get_ngx_api_data("topsymbols")
losers_df = get_ngx_api_data("bottomsymbols")

with col1:
    st.success("### 📈 Top Gainers")
    if not gainers_df.empty:
        st.dataframe(gainers_df, use_container_width=True, hide_index=True)

with col2:
    st.error("### 📉 Top Losers")
    if not losers_df.empty:
        st.dataframe(losers_df, use_container_width=True, hide_index=True)

# 5. Download Section
st.divider()
st.subheader("📁 Export Reports")
btn_col1, btn_col2 = st.columns(2)

# --- EXCEL DOWNLOAD ---
if not gainers_df.empty or not losers_df.empty:
    excel_bio = io.BytesIO()
    with pd.ExcelWriter(excel_bio, engine='openpyxl') as writer:
        if not gainers_df.empty:
            gainers_df.to_excel(writer, sheet_name='Top Gainers', index=False)
        if not losers_df.empty:
            losers_df.to_excel(writer, sheet_name='Top Losers', index=False)
    excel_bio.seek(0)
    with btn_col1:
        st.download_button(label="📊 Download Excel Report", data=excel_bio.getvalue(), file_name=f"NGX_Report_{get_wat_time().strftime('%Y-%m-%d')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

# --- WORD DOWNLOAD ---
if not gainers_df.empty or not losers_df.empty:
    doc = Document()
    doc.add_heading(f'NGX Market Summary', 0)
    doc.add_paragraph(f"Generated on: {get_wat_time().strftime('%d %b %Y at %I:%M:%S %p')} WAT")
    
    # (Word table logic remains the same as previous version)
    # ... code omitted for brevity but should be kept in your file ...

    word_bio = io.BytesIO()
    doc.save(word_bio)
    with btn_col2:
        st.download_button(label="📝 Download Word Report", data=word_bio.getvalue(), file_name=f"NGX_Report_{get_wat_time().strftime('%Y-%m-%d')}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)

# 6. LIVE CLOCK LOOP (Runs at the bottom to keep the UI responsive)
while True:
    now_wat = get_wat_time()
    clock_placeholder.markdown(f"📅 Date: **{now_wat.strftime('%d %b %Y')}** | 🕒 Time: **{now_wat.strftime('%I:%M:%S %p')} WAT**")
    time.sleep(1)
