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

# Word Download Button
if not gainers_df.empty:
    doc = Document()
    doc.add_heading('NGX Market Report', 0)
    # (Your existing Word logic here...)
    bio = io.BytesIO()
    doc.save(bio)
    
    st.download_button(
        label="Download Word Report",
        data=bio.getvalue(),
        file_name=f"NGX_Report_{datetime.date.today()}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
