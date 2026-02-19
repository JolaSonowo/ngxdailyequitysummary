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
        st.dataframe(losers_df, use_container_
