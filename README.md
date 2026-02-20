## NGX Live Equity Summary Cloud Version
This is a professional real-time dashboard that tracks the Top 5 Advancers and Decliners on the Nigerian Exchange Group (NGX). 


It is hosted on Streamlit Cloud for global access.

## Features
Live WAT Clock: Displays the current time in West Africa Time with seconds.

Real-time Data: Fetches the latest market movement directly from the NGX API.

Excel + Word Reports: Download reports in Excel (.xlsx) and Word (.docx) formats.

Auto-Update: Data is cached and refreshed every 60 seconds to ensure accuracy without server lag.

## Setup for Developers
1. Clone the repo: `git clone <git@github.com:JolaSonowo/ngx-market-reports.git>`
2. Install dependencies: `pip install -r requirements.txt`
3. Run the app: `streamlit run ngx_app.py`

## Dependencies
`streamlit`, `pandas`, `requests`, `python-docx`, `openpyxl`, `pytz`

https://ngxsummary.streamlit.app
