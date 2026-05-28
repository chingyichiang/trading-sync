"""
sync_trades.py
--------------
Reads a 群益期貨 海外期貨歷史成交 Excel export, applies LIFO matching
for 微型歐元 trades, and syncs the result to a Google Sheet.

Key design:
- Open buy stack is persisted in a hidden "_state" sheet between runs
- Each open buy records its sheet row number so when a sell arrives later,
  the existing 存貨 row is updated in-place rather than duplicated
- Used keys are persisted to prevent duplicate imports

Usage:
    python sync_trades.py <path_to_export.xlsx>

Requirements:
    pip install pandas openpyxl gspread google-auth
"""

import sys
import json
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime, timedelta

# ── Configuration ──────────────────────────────────────────────────────────────

SPREADSHEET_ID   = "[sheet_id]" # change to your target Google Sheet id
SHEET_NAME       = "[sheet_name]"   # change to your target sheet tab name
STATE_SHEET_NAME = "_state"
CREDENTIALS_FILE = "credentials.json"
PRODUCT_FILTER   = "微型歐元"

# Column indices (0-based)
# A         B        C    D           E         F        G          H        I     J       K          L
# buy_date  b_price  fee  unreal_pnl  tgt_sell  buy_key  sell_date  s_price  s_fee profit  p_rate     sell_key
COL_BUY_DATE    = 0
COL_BUY_PRICE   = 1
COL_BUY_FEE     = 2
COL_UNREAL_PNL  = 3
COL_TARGET_SELL = 4
COL_BUY_KEY     = 5   # F
COL_SELL_DATE   = 6   # G
COL_SELL_PRICE  = 7   # H
COL_SELL_FEE    = 8   # I
COL_PROFIT      = 9   # J
COL_PROFIT_RATE = 10  # K
COL_SELL_KEY    = 11  # L

DATA_START_ROW = 3

# ── Helpers ────────────────────────────────────────────────────────────────────

def parse_trade_time(raw: str) -> datetime:
    raw = str(raw).strip()
    date_part, time_part = raw.split(" ")
    h, m, s = map(int, time_part.split(":"))
    base = datetime.strptime(date_part, "%Y%m%d")
    return base + timedelta(hours=h, minutes=m, seconds=s)

def trade_date(raw: str) -> datetime:
    """Use the date part only — never roll over due to hour overflow."""
    date_part = str(raw).strip().split(" ")[0]
    return datetime.strptime(date_part, "%Y%m%d")

def build_key(order_no: str, price: float, unit_index: int) -> str:
    return f"{order_no}_{price}_{unit_index}"

def excel_date(dt: datetime) -> float:
    epoch = datetime(1899, 12, 30)
    delta = dt - epoch
    return delta.days + delta.seconds / 86400

# ── Load & expand trades ───────────────────────────────────────────────────────

def load_trades(filepath: str) -> list[dict]:
    df = pd.read_excel(filepath, dtype={"委託書號": str})
    df = df[df["商品名稱"] == PRODUCT_FILTER].copy()

    units = []
    for _, row in df.iterrows():
        qty      = int(row["成交量"])
        price    = float(row["成交價"])
        order_no = str(row["委託書號"]).strip()
        time_raw = str(row["成交時間"]).strip()
        side     = str(row["買賣別"]).strip()
        for i in range(1, qty + 1):
            units.append({
                "dt":       parse_trade_time(time_raw),
                "date":     trade_date(time_raw),
                "side":     side,
                "price":    price,
                "key":      build_key(order_no, price, i),
                "time_raw": time_raw,
            })

    units.sort(key=lambda x: x["dt"])
    return units

# ── State persistence ──────────────────────────────────────────────────────────

def get_or_create_state_sheet(spreadsheet) -> gspread.Worksheet:
    try:
        return spreadsheet.worksheet(STATE_SHEET_NAME)
    except gspread.WorksheetNotFound:
        ws = spreadsheet.add_worksheet(title=STATE_SHEET_NAME, rows=10, cols=2)
        ws.update("A1:A2", [["stack_json"], ["used_keys_json"]])
        print(f"  Created state sheet '{STATE_SHEET_NAME}'")
        return ws

def load_state(state_ws) -> tuple[list[dict], set[str]]:
    """
    Load persisted stack and used keys.
    Stack entries include 'sheet_row' — the row they were written to in the
    trading sheet — so we can update them in-place when a sell arrives.
    """
    vals = state_ws.col_values(2)  # column B
    stack     = json.loads(vals[0]) if len(vals) > 0 and vals[0] else []
    used_keys = set(json.loads(vals[1])) if len(vals) > 1 and vals[1] else set()

    for entry in stack:
        entry["dt"]   = datetime.fromisoformat(entry["dt"])
        entry["date"] = datetime.fromisoformat(entry["date"])

    return stack, used_keys

def save_state(state_ws, stack: list[dict], used_keys: set[str]):
    serialisable = []
    for entry in stack:
        e = dict(entry)
        e["dt"]   = entry["dt"].isoformat()
        e["date"] = entry["date"].isoformat()
        serialisable.append(e)

    state_ws.update("B1:B2", [
        [json.dumps(serialisable)],
        [json.dumps(list(used_keys))],
    ])

# ── LIFO matching ──────────────────────────────────────────────────────────────

def lifo_match(new_trades: list[dict], stack: list[dict], used_keys: set[str]) \
        -> tuple[list[dict], list[dict], list[dict]]:
    """
    Process new trades against the carry-over stack.

    Returns:
        new_rows     — buy+存貨 rows to append to the sheet (no sheet_row yet)
        updates      — list of {sheet_row, sell_date, sell_price, sell_key}
                       for existing 存貨 rows that now have a matching sell
        stack        — updated open buy stack (each entry has 'sheet_row' once written)
    """
    new_rows = []   # buys introduced this run (to be appended)
    updates  = []   # existing 存貨 rows to patch with sell details
    buy_order = 0

    for t in new_trades:
        if t["key"] in used_keys:
            continue
        used_keys.add(t["key"])

        if t["side"] == "買進":
            t["_order"]    = buy_order
            t["sheet_row"] = None   # assigned later when written to sheet
            buy_order += 1
            stack.append(t)

        else:  # 賣出
            if stack:
                buy = stack.pop()   # LIFO

                if buy.get("sheet_row") is not None:
                    # This buy was written in a previous run as 存貨 → update in place
                    updates.append({
                        "sheet_row":  buy["sheet_row"],
                        "sell_date":  t["date"],
                        "sell_price": t["price"],
                        "sell_key":   t["key"],
                    })
                else:
                    # Buy was introduced this run → emit as a matched row
                    new_rows.append({
                        "buy_date":   buy["date"],
                        "buy_price":  buy["price"],
                        "sell_date":  t["date"],
                        "sell_price": t["price"],
                        "buy_key":    buy["key"],
                        "sell_key":   t["key"],
                        "_order":     buy["_order"],
                    })
            else:
                print(f"  ⚠️  Sell with no open buy — skipped: {t['time_raw']} price={t['price']}")

    # Remaining new buys still in stack → emit as 存貨 rows
    for entry in stack:
        if entry.get("sheet_row") is None:   # introduced this run
            new_rows.append({
                "buy_date":   entry["date"],
                "buy_price":  entry["price"],
                "sell_date":  None,
                "sell_price": None,
                "buy_key":    entry["key"],
                "sell_key":   "",
                "_order":     entry["_order"],
            })

    new_rows.sort(key=lambda x: x["_order"])
    return new_rows, updates, stack

# ── Google Sheets sync ─────────────────────────────────────────────────────────

def get_spreadsheet():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=scopes)
    return gspread.authorize(creds).open_by_key(SPREADSHEET_ID)

def find_next_empty_row(ws) -> int:
    col_a = ws.col_values(COL_BUY_DATE + 1)
    for i in range(DATA_START_ROW - 1, len(col_a)):
        if not col_a[i]:
            return i + 1
    return len(col_a) + 1

def buy_row_values(trade_row: dict, sheet_row: int) -> list:
    """Full row with buy side filled; sell side empty (存貨) or filled (matched)."""
    r     = sheet_row
    cells = [""] * 12

    cells[COL_BUY_DATE]    = excel_date(trade_row["buy_date"])
    cells[COL_BUY_PRICE]   = trade_row["buy_price"]
    cells[COL_BUY_FEE]     = "=$O$26"
    cells[COL_UNREAL_PNL]  = f'=if(H{r}=0,($O$13-B{r})*1.25,"")'
    cells[COL_TARGET_SELL] = f'=if(H{r}=0,B{r}+$O$23,"")'
    cells[COL_BUY_KEY]     = trade_row["buy_key"]

    if trade_row["sell_date"] is not None:
        cells[COL_SELL_DATE]   = excel_date(trade_row["sell_date"])
        cells[COL_SELL_PRICE]  = trade_row["sell_price"]
        cells[COL_SELL_FEE]    = "=$O$26"
        cells[COL_PROFIT]      = f"=(H{r}-B{r})*$O$27-(C{r}+I{r})"
        cells[COL_PROFIT_RATE] = f"=J{r}/$O$28"

    cells[COL_SELL_KEY] = trade_row["sell_key"]
    return cells

def sell_patch_values(update: dict, sheet_row: int) -> list:
    """Only the sell-side cells (G–L) to patch an existing 存貨 row."""
    r = sheet_row
    return [
        excel_date(update["sell_date"]),   # G
        update["sell_price"],              # H
        "=$O$26",                          # I fee
        f"=(H{r}-B{r})*$O$27-(C{r}+I{r})",  # J profit
        f"=J{r}/$O$28",                    # K profit rate
        update["sell_key"],                # L
    ]

def sync_to_sheet(new_rows: list[dict], updates: list[dict], ws) -> dict[str, int]:
    """
    Append new rows and patch existing 存貨 rows.
    Returns a map of buy_key → sheet_row for newly written rows.
    """
    key_to_row = {}

    # 1. Append new rows
    if new_rows:
        start_row = find_next_empty_row(ws)
        print(f"Appending {len(new_rows)} new row(s) from row {start_row}...")
        batch = []
        for i, row in enumerate(new_rows):
            sr = start_row + i
            batch.append(buy_row_values(row, sr))
            key_to_row[row["buy_key"]] = sr
        ws.update(f"A{start_row}:L{start_row + len(batch) - 1}",
                  batch, value_input_option="USER_ENTERED")
        print(f"  Done. Rows {start_row}–{start_row + len(batch) - 1} written.")
    else:
        print("No new rows to append.")

    # 2. Patch existing 存貨 rows with sell details
    if updates:
        print(f"Updating {len(updates)} existing 存貨 row(s) with sell details...")
        for upd in updates:
            sr = upd["sheet_row"]
            ws.update(f"G{sr}:L{sr}",
                      [sell_patch_values(upd, sr)],
                      value_input_option="USER_ENTERED")
            print(f"  Row {sr} updated with sell @ {upd['sell_price']}")
    else:
        print("No existing 存貨 rows to update.")

    return key_to_row

# ── Main ───────────────────────────────────────────────────────────────────────

def main():
    if len(sys.argv) < 2:
        print("Usage: python sync_trades.py <path_to_export.xlsx>")
        sys.exit(1)

    filepath = sys.argv[1]
    print(f"Loading trades from: {filepath}")
    trades = load_trades(filepath)
    print(f"  {sum(1 for t in trades if t['side']=='買進')} buys, "
          f"{sum(1 for t in trades if t['side']=='賣出')} sells loaded")

    print("Connecting to Google Sheets...")
    spreadsheet = get_spreadsheet()
    ws          = spreadsheet.worksheet(SHEET_NAME)
    state_ws    = get_or_create_state_sheet(spreadsheet)

    print("Loading saved stack state...")
    stack, used_keys = load_state(state_ws)
    print(f"  {len(stack)} open positions carried over, "
          f"{len(used_keys)} keys already processed")

    new_rows, updates, updated_stack = lifo_match(trades, stack, used_keys)

    matched   = sum(1 for r in new_rows if r["sell_date"])
    inventory = sum(1 for r in new_rows if not r["sell_date"])
    print(f"  {matched} new matched pairs  |  "
          f"{inventory} new 存貨 rows  |  "
          f"{len(updates)} existing 存貨 rows to fill in")

    key_to_row = sync_to_sheet(new_rows, updates, ws)

    # Assign sheet_row back to stack entries that were just written
    for entry in updated_stack:
        if entry.get("sheet_row") is None and entry["key"] in key_to_row:
            entry["sheet_row"] = key_to_row[entry["key"]]

    print("Saving stack state for next run...")
    save_state(state_ws, updated_stack, used_keys)
    print(f"  {len(updated_stack)} open positions saved to state.")

if __name__ == "__main__":
    main()
