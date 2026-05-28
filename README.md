# 群益期貨 → Google Sheets Trade Sync

Automatically syncs overseas futures trade history from 群益期貨 (Capital Securities)
exports into a Google Sheet with LIFO matching.

## Features
- LIFO buy/sell matching per position
- Handles partial fills (qty > 1 expanded into individual unit rows)
- Incremental sync — only imports new trades, skips duplicates
- Persists open buy stack between runs so sells in later files correctly
  match buys from earlier files
- Updates existing 存貨 rows in-place when a matching sell arrives

## Setup
1. Install dependencies:
   pip install pandas openpyxl gspread google-auth python-dotenv

2. Set up Google credentials:
   - Enable Google Sheets API and Google Drive API in Google Cloud Console
   - Create a Service Account and download the JSON key as credentials.json
   - Share your Google Sheet with the service account email (Editor access)

3. Copy .env.example to .env and fill in your values:
   cp .env.example .env

4. Run:
   python sync_trades.py path/to/your_export.xlsx
