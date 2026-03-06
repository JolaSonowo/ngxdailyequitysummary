import streamlit as st
import datetime
import requests
import pandas as pd
import io
import pytz
from docx import Document

TOP_N = 7

st.set_page_config(page_title="NGX Market Dashboard", layout="wide")

def get_wat_time():
    utc_now = datetime.datetime.now(datetime.timezone.utc)
    wat_tz = pytz.timezone("Africa/Lagos")
    return utc_now.astimezone(wat_tz)

@st.cache_data(ttl=60)
def get_ngx_api_data(endpoint):
    url = f"https://doclib.ngxgroup.com/REST/api/mrkstat/{endpoint}"
    headers = {
        "User-Agent": "Mozilla/5.0",
        "Referer": "https://ngxgroup.com/"
    }

    response = requests.get(url, headers=headers, timeout=10)
    response.raise_for_status()
    data = response.json()

    return data

st.title("NGX Daily Equity Summary")
now = get_wat_time()
st.markdown(f"**{now.strftime('%d %b %Y')}** | **{now.strftime('%I:%M:%S %p')} WAT**")

raw_gainers = get_ngx_api_data("topsymbols")
raw_losers = get_ngx_api_data("bottomsymbols")

st.write("Gainers rows returned by API:", len(raw_gainers))
st.write("Losers rows returned by API:", len(raw_losers))

def format_data(data):
    output = []
    for item in data[:TOP_N]:
        last_close = float(item.get("LAST_CLOSE", 0) or 0)
        percentage_change = float(item.get("PERCENTAGE_CHANGE", 0) or 0)
        todays_close = float(item.get("TODAYS_CLOSE", 0) or 0)

        output.append({
            "Symbol": item.get("SYMBOL", "N/A"),
            "Price": todays_close,
            "Change (N)": percentage_change,
            "% Change": round((percentage_change / last_close * 100), 2) if last_close != 0 else 0
        })
    return pd.DataFrame(output)

gainers_df = format_data(raw_gainers)
losers_df = format_data(raw_losers)

col1, col2 = st.columns(2)

with col1:
    st.success(f"**Top {TOP_N} Advancers**")
    st.dataframe(gainers_df, use_container_width=True, hide_index=True)

with col2:
    st.error(f"**Top {TOP_N} Decliners**")
    st.dataframe(losers_df, use_container_width=True, hide_index=True)            
    last_close = float(item.get("LAST_CLOSE", 0) or 0)
            percentage_change = float(item.get("PERCENTAGE_CHANGE", 0) or 0)
            todays_close = float(item.get("TODAYS_CLOSE", 0) or 0)

            output.append({
                "Symbol": item.get("SYMBOL", "N/A"),
                "Price": todays_close,
                "Change (N)": percentage_change,
                "% Change": round((percentage_change / last_close * 100), 2) if last_close != 0 else 0
            })

        return pd.DataFrame(output)

    except Exception as e:
        st.error(f"Error fetching {endpoint}: {e}")
        return pd.DataFrame()

# 3. HEADER & CLOCK
st.title("NGX Daily Equity Summary")

now = get_wat_time()
st.markdown(f"**{now.strftime('%d %b %Y')}** | **{now.strftime('%I:%M:%S %p')} WAT**")

# 4. MAIN DASHBOARD CONTENT
col1, col2 = st.columns(2)

gainers_df = get_ngx_api_data("topsymbols")
losers_df = get_ngx_api_data("bottomsymbols")

with col1:
    st.success(f"**Top {TOP_N} Advancers**")
    if not gainers_df.empty:
        st.dataframe(gainers_df, use_container_width=True, hide_index=True)
    else:
        st.warning("No gainers data available currently.")

with col2:
    st.error(f"**Top {TOP_N} Decliners**")
    if not losers_df.empty:
        st.dataframe(losers_df, use_container_width=True, hide_index=True)
    else:
        st.warning("No losers data available currently.")

# 5. DOWNLOAD SECTION
st.divider()
st.subheader("Export Reports")

btn_col1, btn_col2 = st.columns(2)

if not gainers_df.empty or not losers_df.empty:
    # Excel download
    excel_bio = io.BytesIO()
    with pd.ExcelWriter(excel_bio, engine="openpyxl") as writer:
        if not gainers_df.empty:
            gainers_df.to_excel(writer, sheet_name="Top Gainers", index=False)
        if not losers_df.empty:
            losers_df.to_excel(writer, sheet_name="Top Losers", index=False)

    excel_bio.seek(0)

    with btn_col1:
        st.download_button(
            label="Download Excel Report",
            data=excel_bio.getvalue(),
            file_name=f"NGX_Report_{get_wat_time().strftime('%Y-%m-%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

    # Word download
    doc = Document()
    current_wat = get_wat_time()
    doc.add_heading("NGX Market Summary", 0)
    doc.add_paragraph(f"Generated on: {current_wat.strftime('%d %b %Y at %I:%M:%S %p')} WAT")

    if not gainers_df.empty:
        doc.add_heading(f"Top {TOP_N} Gainers", level=1)
        table = doc.add_table(rows=1, cols=len(gainers_df.columns))
        table.style = "Table Grid"

        for i, column in enumerate(gainers_df.columns):
            table.rows[0].cells[i].text = str(column)

        for _, row in gainers_df.iterrows():
            row_cells = table.add_row().cells
            for i, value in enumerate(row):
                row_cells[i].text = str(value)

    if not losers_df.empty:
        doc.add_heading(f"Top {TOP_N} Losers", level=1)
        table_l = doc.add_table(rows=1, cols=len(losers_df.columns))
        table_l.style = "Table Grid"

        for i, column in enumerate(losers_df.columns):
            table_l.rows[0].cells[i].text = str(column)

        for _, row in losers_df.iterrows():
            row_cells = table_l.add_row().cells
            for i, value in enumerate(row):
                row_cells[i].text = str(value)

    word_bio = io.BytesIO()
    doc.save(word_bio)
    word_bio.seek(0)

    with btn_col2:
        st.download_button(
            label="Download Word Report",
            data=word_bio.getvalue(),
            file_name=f"NGX_Report_{current_wat.strftime('%Y-%m-%d')}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True
        )
