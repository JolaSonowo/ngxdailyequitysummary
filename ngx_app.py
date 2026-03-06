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

# 2. DATA FETCHING 
@st.cache_data(ttl=60)
def get_ngx_api_data(endpoint):
    url = f"https://doclib.ngxgroup.com/REST/api/mrkstat/{endpoint}"
    headers = {"User-Agent": "Mozilla/5.0", "Referer": "https://ngxgroup.com/"}
    try:
        response = requests.get(url, headers=headers, timeout=10)
        data = response.json()
        output = []
        for item in data[:7]:
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

clock_placeholder = st.empty()

# 4. MAIN DASHBOARD CONTENT
col1, col2 = st.columns(2)


gainers_df = get_ngx_api_data("topsymbols")
losers_df = get_ngx_api_data("bottomsymbols")

with col1:
    st.success("**Top 7 Advancers**")
    if not gainers_df.empty:
        st.dataframe(gainers_df, use_container_width=True, hide_index=True)
    else:
        st.warning("No gainers data available currently.")

with col2:
    st.error("**Top 7 Decliners**")
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
from docx.shared import RGBColor, Pt
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# --- HELPERS FOR STYLING ---
def set_cell_background(cell, fill_color):
    shading_elm = OxmlElement('w:shd')
    shading_elm.set(qn('w:fill'), fill_color)
    cell._tc.get_or_add_tcPr().append(shading_elm)

def apply_font_style(cell, font_name, size, is_bold=False, color=None):
    for paragraph in cell.paragraphs:
        for run in paragraph.runs:
            run.font.name = font_name
            run.font.size = Pt(size)
            run.bold = is_bold
            if color:
                run.font.color.rgb = RGBColor.from_string(color)
            # Necessary for Word to recognize the font name correctly
            r = run._element
            r.rPr.rFonts.set(qn('w:ascii'), font_name)
            r.rPr.rFonts.set(qn('w:hAnsi'), font_name)

# --- UPDATED WORD DOWNLOAD LOGIC ---
from docx.shared import RGBColor, Pt
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# --- STYLING HELPERS ---
def set_cell_background(cell, fill_color):
    shading_elm = OxmlElement('w:shd')
    shading_elm.set(qn('w:fill'), fill_color)
    cell._tc.get_or_add_tcPr().append(shading_elm)

def apply_table_font(cell, font_name, size, is_bold=False):
    for paragraph in cell.paragraphs:
        for run in paragraph.runs:
            run.font.name = font_name
            run.font.size = Pt(size)
            run.bold = is_bold
            # Force Word to recognize the font
            r = run._element
            r.rPr.rFonts.set(qn('w:ascii'), font_name)
            r.rPr.rFonts.set(qn('w:hAnsi'), font_name)

# --- WORD GENERATION ---
if not gainers_df.empty or not losers_df.empty:
    doc = Document()
    current_wat = get_wat_time()
    
    # Main Doc Header: Calibri 12
    title = doc.add_heading('NGX Market Summary', 0)
    title.runs[0].font.name = 'Calibri'
    title.runs[0].font.size = Pt(12)

    # 1. TOP GAINERS SECTION
    if not gainers_df.empty:
        gainer_heading = doc.add_heading('Top Gainers', level=1)
        g_run = gainer_heading.runs[0]
        g_run.font.name = 'Calibri'
        g_run.font.size = Pt(12)
        g_run.font.color.rgb = RGBColor(255, 255, 255) # White
        
        table = doc.add_table(rows=1, cols=4)
        table.style = 'Table Grid'
        
        # Header Row: Calibri Bold 11, Green Background
        cols = ["Gainers", "Close Price", "% Change", "Naira Change"]
        for i, text in enumerate(cols):
            cell = table.rows[0].cells[i]
            cell.text = text
            set_cell_background(cell, "548235") # Darker green like image 2
            apply_table_font(cell, "Calibri", 11, is_bold=True)

        # Body Rows: Aptos 11, Green Background
        for _, row in gainers_df.iterrows():
            row_cells = table.add_row().cells
            vals = [row['Symbol'], f"{row['Price']:.2f}", f"{row['% Change']:.2f}", f"{row['Change (N)']:.2f}"]
            for i, val in enumerate(vals):
                row_cells[i].text = str(val)
                set_cell_background(row_cells[i], "548235")
                apply_table_font(row_cells[i], "Aptos", 11)

    doc.add_paragraph("\n") 

    # 2. TOP LOSERS SECTION
    if not losers_df.empty:
        loser_heading = doc.add_heading('Top Losers', level=1)
        l_run = loser_heading.runs[0]
        l_run.font.name = 'Calibri'
        l_run.font.size = Pt(12)
        l_run.font.color.rgb = RGBColor(255, 255, 255) # White
        
        table_l = doc.add_table(rows=1, cols=4)
        table_l.style = 'Table Grid'
        
        # Header Row: Calibri Bold 11, Red Background
        cols_l = ["Losers", "Close Price", "% Change", "Naira Change"]
        for i, text in enumerate(cols_l):
            cell = table_l.rows[0].cells[i]
            cell.text = text
            set_cell_background(cell, "843C39") # Darker red like image 2
            apply_table_font(cell, "Calibri", 11, is_bold=True)

        # Body Rows: Aptos 11, Red Background
        for _, row in losers_df.iterrows():
            row_cells = table_l.add_row().cells
            vals = [row['Symbol'], f"{row['Price']:.2f}", f"{row['% Change']:.2f}", f"{row['Change (N)']:.2f}"]
            for i, val in enumerate(vals):
                row_cells[i].text = str(val)
                set_cell_background(row_cells[i], "843C39")
                apply_table_font(row_cells[i], "Aptos", 11)
                
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


# 6. LIVE CLOCK 
while True:
    now = get_wat_time()
    clock_placeholder.markdown(
        f" **{now.strftime('%d %b %Y')}** | **{now.strftime('%I:%M:%S %p')} WAT**"
    )
    time.sleep(1)
